#![cfg_attr(
    all(not(debug_assertions), target_os = "windows"),
    windows_subsystem = "windows"
)]

use lettre::message::{header::ContentType, Mailbox};
use lettre::transport::smtp::authentication::Credentials;
use lettre::{Message, SmtpTransport, Transport};
use regex::Regex;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::PathBuf;
use tauri::{Emitter, Manager};

const DEFAULT_SUBJECT_TEMPLATE: &str = "【薪酬通知】{{月份}}月工资条 - {{人员}}";
const DEFAULT_BODY_TEMPLATE: &str = r#"<p style="margin: 0 0 12px;">{{人员}}，您好：</p>

<p style="margin: 0 0 12px;">现将您 <strong>{{月份}} 月</strong> 工资发放情况通知如下，请您查收并妥善保管。</p>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; width: auto; max-width: 100%; border-color: #d9dee8; font-size: 12px; line-height: 1.45; margin: 0 0 12px;">
  <tr style="background: #f7f9fc;">
    <th align="left" width="132" style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">项目</th>
    <th align="left" style="padding: 6px 8px; border: 1px solid #d9dee8;">内容</th>
  </tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">备注</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{备注}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">序号</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{序号}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">月份</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{月份}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">公司</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{公司}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">人员</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{人员}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">应 出勤</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{应 出勤}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">实际 出勤</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{实际 出勤}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">二级部门</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{二级部门}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">三级部门</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{三级部门}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">固定基本工资</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{固定基本工资}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">岗位工资</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{岗位工资}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">绩效奖金</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{绩效奖金}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">补贴</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{补贴}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">新增/试用 调整</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{新增/试用 调整}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">离职调整</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{离职调整}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">电脑补贴</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{电脑补贴}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">特殊津贴 月度提成</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{特殊津贴 月度提成}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">派驻补贴</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{派驻补贴}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">差旅补贴</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{差旅补贴}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">离职补偿</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{离职补偿}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">兼职津贴</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{兼职津贴}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">奖金</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{奖金}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">加班津贴</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{加班津贴}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">事假天数</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{事假天数}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">病假扣款</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{病假扣款}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">事假扣款</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{事假扣款}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">迟到早退 未打卡扣款</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{迟到早退 未打卡扣款}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">其他扣款</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{其他扣款}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">补发 或 扣发</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{补发 或 扣发}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">工资总额</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{工资总额}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">养老保险</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{养老保险}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">医疗保险</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{医疗保险}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">失业保险</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{失业保险}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">公积金</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{公积金}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">个人小计</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{个人小计}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">本期累计预扣预缴应税额</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{本期累计预扣预缴应税额}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">适用税率</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{适用税率}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">速算 扣除数</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{速算 扣除数}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">本期累计代扣代缴 个人所得税</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{本期累计代扣代缴 个人所得税}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">往期累计已预扣预缴税额</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{往期累计已预扣预缴税额}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">往来扣款</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{往来扣款}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">其他</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{其他}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">实发工资</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{实发工资}}</td></tr>
  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">本期实际应预扣预缴税额税额</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{本期实际应预扣预缴税额税额}}</td></tr>
</table>

<p style="margin: 0 0 12px;">如您对本人工资条信息有疑问，请在收到邮件后及时与财务或人力资源相关同事联系核实。</p>
<p style="margin: 0 0 12px;">感谢您的配合。</p>
<p style="margin: 0 0 12px;">此致<br/>敬礼</p>
<p style="margin: 0 0 12px;">财务部 / 人力资源部<br/>工资条通知系统</p>
<p style="margin: 0; color: #6b7280; font-size: 12px;">本邮件包含个人薪酬敏感信息，仅限收件人本人查阅，请勿擅自转发、传播或用于其他用途。</p>"#;

#[derive(Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct PayslipSettings {
    default_subject_template: String,
    default_body_template: String,
    smtp: SmtpSettings,
}

#[derive(Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct SmtpSettings {
    host: String,
    port: String,
    username: String,
    password: String,
    from_address: String,
    from_name: String,
    auth: bool,
    starttls: bool,
}

#[derive(Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct SendRequest {
    subject_template: String,
    body_template: String,
    smtp: SmtpSettings,
    rows: Vec<SendRow>,
}

#[derive(Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct SendRow {
    row_number: i32,
    recipient_name: String,
    email: String,
    values: HashMap<String, String>,
}

#[derive(Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SendResponse {
    total_count: usize,
    success_count: usize,
    failure_count: usize,
    results: Vec<SendResult>,
}

#[derive(Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SendResult {
    row_number: i32,
    recipient_name: String,
    email: String,
    status: String,
    message: String,
    values: HashMap<String, String>,
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct SendProgressPayload {
    processed_count: usize,
    total_count: usize,
    row_number: i32,
    recipient_name: String,
    email: String,
    status: String,
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            payslip_get_settings,
            payslip_save_settings,
            payslip_send
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[tauri::command]
fn payslip_get_settings(app: tauri::AppHandle) -> Result<PayslipSettings, String> {
    let path = settings_path(&app)?;
    if !path.exists() {
        return Ok(default_settings());
    }
    let content =
        fs::read_to_string(path).map_err(|error| format!("工资条配置读取失败: {error}"))?;
    let settings: PayslipSettings =
        serde_json::from_str(&content).map_err(|error| format!("工资条配置读取失败: {error}"))?;
    Ok(settings)
}

#[tauri::command]
fn payslip_save_settings(
    app: tauri::AppHandle,
    settings: PayslipSettings,
) -> Result<PayslipSettings, String> {
    validate_smtp_settings_for_save(&settings.smtp)?;
    let path = settings_path(&app)?;
    let content = serde_json::to_string_pretty(&settings)
        .map_err(|error| format!("工资条配置序列化失败: {error}"))?;
    fs::write(path, content).map_err(|error| format!("工资条配置保存失败: {error}"))?;
    Ok(settings)
}

#[tauri::command]
fn payslip_send(app: tauri::AppHandle, request: SendRequest) -> Result<SendResponse, String> {
    if request.rows.is_empty() {
        return Err("至少需要一条待发送记录".into());
    }
    validate_smtp_settings(&request.smtp)?;

    let total_count = request.rows.len();
    let mailer = build_mailer(&request.smtp)?;
    let from_mailbox = Mailbox::new(
        Some(normalize_or_default(
            &request.smtp.from_name,
            "工资条通知",
        )),
        request
            .smtp
            .from_address
            .trim()
            .parse()
            .map_err(|error| format!("发件邮箱格式不正确: {error}"))?,
    );

    let mut results = Vec::with_capacity(total_count);
    let mut success_count = 0usize;

    for (index, row) in request.rows.into_iter().enumerate() {
        let send_result = match send_one_mail(
            &mailer,
            &from_mailbox,
            &request.subject_template,
            &request.body_template,
            &row,
        ) {
            Ok(()) => {
                success_count += 1;
                SendResult {
                    row_number: row.row_number,
                    recipient_name: row.recipient_name,
                    email: row.email,
                    status: "SUCCESS".into(),
                    message: "发送成功".into(),
                    values: row.values,
                }
            }
            Err(message) => {
                SendResult {
                    row_number: row.row_number,
                    recipient_name: row.recipient_name,
                    email: row.email,
                    status: "FAILED".into(),
                    message,
                    values: row.values,
                }
            }
        };

        emit_send_progress(&app, &send_result, index + 1, total_count)?;
        results.push(send_result);
    }

    append_history(
        &app,
        &format!(
            "{} success={} failure={}",
            chrono::Local::now().to_rfc3339(),
            success_count,
            results.len() - success_count
        ),
    )?;

    Ok(SendResponse {
        total_count: results.len(),
        success_count,
        failure_count: results.len() - success_count,
        results,
    })
}

fn emit_send_progress(
    app: &tauri::AppHandle,
    result: &SendResult,
    processed_count: usize,
    total_count: usize,
) -> Result<(), String> {
    app.emit(
        "payslip-send-progress",
        SendProgressPayload {
            processed_count,
            total_count,
            row_number: result.row_number,
            recipient_name: result.recipient_name.clone(),
            email: result.email.clone(),
            status: result.status.clone(),
        },
    )
    .map_err(|error| format!("无法发送进度事件: {error}"))
}

fn send_one_mail(
    mailer: &SmtpTransport,
    from_mailbox: &Mailbox,
    subject_template: &str,
    body_template: &str,
    row: &SendRow,
) -> Result<(), String> {
    if row.email.trim().is_empty() {
        return Err("邮箱为空，无法发送".into());
    }

    let to_mailbox = row
        .email
        .trim()
        .parse()
        .map_err(|error| format!("邮箱格式不正确: {error}"))?;
    let subject = render_template(subject_template, &row.values, false).trim().to_string();
    let body = render_template(body_template, &row.values, true);

    let message = Message::builder()
        .from(from_mailbox.clone())
        .to(to_mailbox)
        .subject(subject)
        .header(ContentType::TEXT_HTML)
        .body(body)
        .map_err(|error| simplify_error(format!("邮件组装失败: {error}")))?;

    mailer
        .send(&message)
        .map_err(|error| simplify_error(format!("{error}")))?;
    Ok(())
}

fn build_mailer(smtp: &SmtpSettings) -> Result<SmtpTransport, String> {
    let host = smtp.host.trim();
    if host.is_empty() {
        return Err("SMTP 主机不能为空".into());
    }

    let port = smtp.port.trim().parse::<u16>().unwrap_or(587);
    let mut builder = if smtp.starttls {
        SmtpTransport::starttls_relay(host)
            .map_err(|error| format!("SMTP 配置无效: {error}"))?
            .port(port)
    } else if port == 465 {
        SmtpTransport::relay(host)
            .map_err(|error| format!("SMTP 配置无效: {error}"))?
            .port(port)
    } else {
        SmtpTransport::builder_dangerous(host).port(port)
    };

    if smtp.auth {
        builder = builder.credentials(Credentials::new(
            smtp.username.trim().to_string(),
            smtp.password.clone(),
        ));
    }

    Ok(builder.build())
}

fn render_template(
    template: &str,
    values: &HashMap<String, String>,
    escape_html: bool,
) -> String {
    let merged = merged_values(values);
    let pattern = Regex::new(r"\{\{\s*(.+?)\s*\}\}").expect("template regex should be valid");
    pattern
        .replace_all(template, |captures: &regex::Captures| {
            let key = normalize_text(captures.get(1).map(|item| item.as_str()).unwrap_or_default());
            let value = merged.get(&key).cloned().unwrap_or_default();
            if escape_html {
                v_htmlescape::escape(&value).to_string()
            } else {
                value
            }
        })
        .to_string()
}

fn merged_values(values: &HashMap<String, String>) -> HashMap<String, String> {
    let mut merged = values.clone();

    for (key, value) in values {
        let compact_key = compact_key(key);
        if compact_key != *key {
            merged.entry(compact_key).or_insert_with(|| value.clone());
        }
    }

    merged
        .entry("姓名".into())
        .or_insert_with(|| value_of(values, "人员"));
    merged
        .entry("邮箱地址".into())
        .or_insert_with(|| value_of(values, "邮箱"));
    merged
        .entry("净发工资".into())
        .or_insert_with(|| value_of(values, "实发工资"));
    merged
}

fn compact_key(value: &str) -> String {
    normalize_text(value).replace(' ', "")
}

fn value_of(values: &HashMap<String, String>, key: &str) -> String {
    values
        .get(key)
        .map(|value| normalize_text(value))
        .unwrap_or_default()
}

fn validate_smtp_settings(smtp: &SmtpSettings) -> Result<(), String> {
    if smtp.host.trim().is_empty() {
        return Err("SMTP 主机不能为空".into());
    }
    if smtp.from_address.trim().is_empty() {
        return Err("发件邮箱不能为空".into());
    }
    if smtp.auth && smtp.username.trim().is_empty() {
        return Err("启用 SMTP 认证时，用户名不能为空".into());
    }
    if smtp.auth && smtp.password.trim().is_empty() {
        return Err("启用 SMTP 认证时，密码不能为空".into());
    }
    Ok(())
}

fn validate_smtp_settings_for_save(smtp: &SmtpSettings) -> Result<(), String> {
    if smtp.host.trim().is_empty() {
        return Err("SMTP 主机不能为空".into());
    }
    if smtp.from_address.trim().is_empty() {
        return Err("发件邮箱不能为空".into());
    }
    Ok(())
}

fn normalize_or_default(value: &str, default_value: &str) -> String {
    let normalized = normalize_text(value);
    if normalized.is_empty() {
        default_value.to_string()
    } else {
        normalized
    }
}

fn normalize_text(value: &str) -> String {
    value.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn simplify_error(message: String) -> String {
    if message.is_empty() {
        return "发送失败，请检查邮件配置或收件箱地址".into();
    }
    message.chars().take(120).collect()
}

fn settings_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    let mut dir = app
        .path()
        .app_config_dir()
        .map_err(|error| format!("无法获取配置目录: {error}"))?;
    dir.push("payslip-mailer");
    fs::create_dir_all(&dir).map_err(|error| format!("无法创建配置目录: {error}"))?;
    dir.push("settings.json");
    Ok(dir)
}

fn history_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    let mut dir = app
        .path()
        .app_config_dir()
        .map_err(|error| format!("无法获取配置目录: {error}"))?;
    dir.push("payslip-mailer");
    fs::create_dir_all(&dir).map_err(|error| format!("无法创建配置目录: {error}"))?;
    dir.push("history.log");
    Ok(dir)
}

fn append_history(app: &tauri::AppHandle, entry: &str) -> Result<(), String> {
    let path = history_path(app)?;
    let mut file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
        .map_err(|error| format!("无法写入发送记录: {error}"))?;
    writeln!(file, "{entry}").map_err(|error| format!("无法追加发送记录: {error}"))?;
    Ok(())
}

fn default_settings() -> PayslipSettings {
    PayslipSettings {
        default_subject_template: DEFAULT_SUBJECT_TEMPLATE.into(),
        default_body_template: DEFAULT_BODY_TEMPLATE.into(),
        smtp: SmtpSettings {
            host: "smtp.mxhichina.com".into(),
            port: "465".into(),
            username: String::new(),
            password: String::new(),
            from_address: String::new(),
            from_name: "工资条通知".into(),
            auth: true,
            starttls: false,
        },
    }
}
