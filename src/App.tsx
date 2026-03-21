import { useEffect, useMemo, useRef, useState } from "react";
import {
  buildDefaultBodyTemplate,
  DEFAULT_SMTP,
  DEFAULT_SUBJECT_TEMPLATE,
  exportFailedRows,
  type PayslipSettings,
  parsePayslipFile,
  type PreviewResponse,
  type SendRequest,
  type SendResponse,
  type SmtpSettings,
} from "./payslip";
import { loadSettings, saveSettings, sendPayslips } from "./runtime";

const PREVIEW_STATUS_TEXT = {
  READY: "可发送",
  INVALID: "校验失败",
  SKIPPED: "已跳过",
} as const;

const SEND_STATUS_TEXT = {
  SUCCESS: "发送成功",
  FAILED: "发送失败",
} as const;

function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [file, setFile] = useState<File | null>(null);
  const [defaultSubjectTemplate, setDefaultSubjectTemplate] = useState(
    DEFAULT_SUBJECT_TEMPLATE,
  );
  const [defaultBodyTemplate, setDefaultBodyTemplate] = useState(
    buildDefaultBodyTemplate(),
  );
  const [defaultSmtp, setDefaultSmtp] = useState<SmtpSettings>(DEFAULT_SMTP);
  const [subjectTemplate, setSubjectTemplate] = useState(
    DEFAULT_SUBJECT_TEMPLATE,
  );
  const [bodyTemplate, setBodyTemplate] = useState(buildDefaultBodyTemplate());
  const [smtp, setSmtp] = useState<SmtpSettings>(DEFAULT_SMTP);
  const [preview, setPreview] = useState<PreviewResponse | null>(null);
  const [sendResult, setSendResult] = useState<SendResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [settingsMessage, setSettingsMessage] = useState<string | null>(null);
  const [isPreviewing, setIsPreviewing] = useState(false);
  const [isSending, setIsSending] = useState(false);
  const [isSavingSettings, setIsSavingSettings] = useState(false);
  const [sendProgress, setSendProgress] = useState(0);

  useEffect(() => {
    let active = true;
    loadSettings()
      .then((settings) => {
        if (!active) {
          return;
        }
        setDefaultSubjectTemplate(settings.defaultSubjectTemplate);
        setDefaultBodyTemplate(settings.defaultBodyTemplate);
        setDefaultSmtp(settings.smtp);
        setSubjectTemplate(settings.defaultSubjectTemplate);
        setBodyTemplate(settings.defaultBodyTemplate);
        setSmtp({
          host: settings.smtp.host,
          port: settings.smtp.port,
          username: settings.smtp.username,
          password: settings.smtp.password,
          fromAddress: settings.smtp.fromAddress,
          fromName: settings.smtp.fromName,
          auth: settings.smtp.auth,
          starttls: settings.smtp.starttls,
        });
      })
      .catch(() => undefined);
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!isSending) {
      setSendProgress(0);
      return;
    }

    setSendProgress(8);
    const timer = window.setInterval(() => {
      setSendProgress((current) => {
        if (current >= 92) {
          return current;
        }
        const next =
          current < 24 ? current + 8 : current < 56 ? current + 5 : current + 3;
        return Math.min(next, 92);
      });
    }, 450);

    return () => {
      window.clearInterval(timer);
    };
  }, [isSending]);

  const sendProgressLabel = useMemo(() => {
    if (!isSending) {
      return null;
    }
    if (sendProgress < 20) {
      return "正在连接 SMTP 服务器";
    }
    if (sendProgress < 55) {
      return "正在认证并建立加密连接";
    }
    return "正在逐条发送工资条邮件";
  }, [isSending, sendProgress]);

  const readyRows = useMemo(
    () =>
      preview?.rows
        .filter((row) => row.status === "READY")
        .map((row) => ({
          rowNumber: row.rowNumber,
          recipientName: row.recipientName,
          email: row.email,
          values: row.values,
        })) ?? [],
    [preview],
  );

  const failedRows = useMemo(() => {
    const invalidPreviewRows =
      preview?.rows
        .filter((row) => row.status === "INVALID")
        .map((row) => ({
          rowNumber: row.rowNumber,
          recipientName: row.recipientName,
          email: row.email,
          message: row.message,
          values: row.values,
        })) ?? [];

    const failedSendRows =
      sendResult?.results
        .filter((row) => row.status === "FAILED")
        .map((row) => ({
          rowNumber: row.rowNumber,
          recipientName: row.recipientName,
          email: row.email,
          message: row.message,
          values: row.values,
        })) ?? [];

    const merged = new Map<number, (typeof invalidPreviewRows)[number]>();
    [...invalidPreviewRows, ...failedSendRows].forEach((row) =>
      merged.set(row.rowNumber, row),
    );
    return Array.from(merged.values()).sort(
      (left, right) => left.rowNumber - right.rowNumber,
    );
  }, [preview, sendResult]);

  async function handlePreview(nextFile?: File | null) {
    const targetFile = nextFile ?? file;
    if (!targetFile) {
      setError("请先选择工资条文件");
      return;
    }
    setError(null);
    setSettingsMessage(null);
    setSendResult(null);
    setIsPreviewing(true);
    try {
      const result = await parsePayslipFile(targetFile);
      setPreview(result);
      if (
        subjectTemplate === DEFAULT_SUBJECT_TEMPLATE &&
        defaultSubjectTemplate === DEFAULT_SUBJECT_TEMPLATE
      ) {
        setSubjectTemplate(result.defaultSubjectTemplate);
      }
      if (
        bodyTemplate === buildDefaultBodyTemplate() &&
        defaultBodyTemplate === buildDefaultBodyTemplate()
      ) {
        setBodyTemplate(result.defaultBodyTemplate);
      }
    } catch (invokeError) {
      setPreview(null);
      setError(
        invokeError instanceof Error ? invokeError.message : "工资条预览失败",
      );
    } finally {
      setIsPreviewing(false);
    }
  }

  async function handleSend() {
    if (readyRows.length === 0) {
      setError("当前没有可发送的工资条记录");
      return;
    }

    setError(null);
    setSettingsMessage(null);
    setIsSending(true);
    try {
      const request: SendRequest = {
        subjectTemplate,
        bodyTemplate,
        smtp,
        rows: readyRows,
      };
      const result = await sendPayslips(request);
      setSendProgress(100);
      setSendResult(result);
    } catch (invokeError) {
      setError(
        invokeError instanceof Error ? invokeError.message : "工资条发送失败",
      );
    } finally {
      setIsSending(false);
    }
  }

  async function handleSaveSettings(nextSettings?: PayslipSettings) {
    const settings = nextSettings ?? {
      defaultSubjectTemplate,
      defaultBodyTemplate,
      smtp,
    };
    setIsSavingSettings(true);
    setError(null);
    setSettingsMessage(null);
    try {
      await saveSettings(settings);
      setSettingsMessage("配置已保存，下次打开应用会自动读取。");
    } catch (invokeError) {
      setError(
        invokeError instanceof Error
          ? invokeError.message
          : "保存工资条配置失败",
      );
    } finally {
      setIsSavingSettings(false);
    }
  }

  function handleRestoreDefaults() {
    setSubjectTemplate(defaultSubjectTemplate);
    setBodyTemplate(defaultBodyTemplate);
  }

  function handleRestoreDefaultSmtp() {
    setSmtp(defaultSmtp);
  }

  function handleResetImportedFile() {
    setFile(null);
    setPreview(null);
    setSendResult(null);
    setError(null);
    setSettingsMessage(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  }

  function handleSyncCurrentAsDefault() {
    const nextSettings = {
      defaultSubjectTemplate: subjectTemplate,
      defaultBodyTemplate: bodyTemplate,
      smtp,
    };
    setDefaultSubjectTemplate(subjectTemplate);
    setDefaultBodyTemplate(bodyTemplate);
    void handleSaveSettings(nextSettings);
  }

  function handleSyncCurrentSmtpAsDefault() {
    const nextSettings = {
      defaultSubjectTemplate,
      defaultBodyTemplate,
      smtp,
    };
    setDefaultSmtp(smtp);
    void handleSaveSettings(nextSettings);
  }

  return (
    <main className="app-shell">
      <section className="hero-card">
        <div className="hero-copy">
          <p className="eyebrow">BeeWorks Desktop Tool</p>
          <h1>工资条发送</h1>
          <p className="hero-text">
            文件解析、模板渲染、SMTP 发信和本地配置都在桌面端独立完成。
          </p>
        </div>
        <div className="hero-metrics">
          <div className="metric">
            <span>待发送</span>
            <strong>{readyRows.length}</strong>
          </div>
          <div className="metric">
            <span>校验失败</span>
            <strong>{preview?.invalidCount ?? 0}</strong>
          </div>
          <div className="metric">
            <span>发送成功</span>
            <strong>{sendResult?.successCount ?? 0}</strong>
          </div>
        </div>
      </section>

      <section className="panel-card">
        <div className="panel-title">
          <div>
            <h2>预览与发送</h2>
          </div>
        </div>
        <div className="panel-body">
          <label className="field field-wide">
            <span>工资条文件</span>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={(event) => {
                const nextFile = event.target.files?.[0] ?? null;
                setFile(nextFile);
                void handlePreview(nextFile);
              }}
            />
          </label>

          <div className="action-row">
            <button
              className="primary-button"
              type="button"
              onClick={handleSend}
              disabled={isSending || readyRows.length === 0}
            >
              {isSending ? "发送中..." : `发送 ${readyRows.length} 条`}
            </button>
            <button
              className="danger-button"
              type="button"
              disabled={!preview || failedRows.length === 0}
              onClick={() => preview && exportFailedRows(preview, failedRows)}
            >
              {`导出预览/发送失败记录 ${failedRows.length} 条`}
            </button>
            <button
              className="ghost-button"
              type="button"
              disabled={!file && !preview && !sendResult}
              onClick={handleResetImportedFile}
            >
              清空当前文件
            </button>
          </div>

          {isSending ? (
            <div className="progress-card" aria-live="polite">
              <div className="progress-copy">
                <strong>发送进行中</strong>
                <span>{sendProgressLabel}</span>
              </div>
              <div className="progress-track">
                <div
                  className="progress-fill"
                  style={{ width: `${sendProgress}%` }}
                />
              </div>
              <p className="helper-text">
                真实 SMTP
                发送会依次完成连接、认证和逐条投递，这段时间页面不是卡住，只是此前缺少过程反馈。
              </p>
            </div>
          ) : null}

          {error ? <p className="feedback error-text">{error}</p> : null}
          {settingsMessage ? (
            <p className="feedback success-text">{settingsMessage}</p>
          ) : null}

          {preview ? (
            <>
              <div className="summary-grid">
                <div className="summary-pill">文件：{preview.fileName}</div>
                <div className="summary-pill">工作表：{preview.sheetName}</div>
                <div className="summary-pill">总记录：{preview.totalCount}</div>
                <div className="summary-pill">可发送：{preview.readyCount}</div>
                <div className="summary-pill">
                  校验失败：{preview.invalidCount}
                </div>
                <div className="summary-pill">
                  已跳过：{preview.skippedCount}
                </div>
              </div>

              <div className="chips">
                {preview.headers
                  .filter(Boolean)
                  .slice(0, 18)
                  .map((header) => (
                    <span key={header} className="chip">{`{{${header}}}`}</span>
                  ))}
              </div>

              <div className="table-shell">
                <table>
                  <thead>
                    <tr>
                      <th>行号</th>
                      <th>人员</th>
                      <th>邮箱</th>
                      <th>月份</th>
                      <th>实发工资</th>
                      <th>状态</th>
                      <th>说明</th>
                    </tr>
                  </thead>
                  <tbody>
                    {preview.rows.map((row) => (
                      <tr
                        key={`${row.rowNumber}-${row.email}-${row.recipientName}`}
                      >
                        <td>{row.rowNumber}</td>
                        <td>{row.recipientName || "-"}</td>
                        <td>{row.email || "-"}</td>
                        <td>{row.month || "-"}</td>
                        <td>{row.netPay || row.values["工资总额"] || "-"}</td>
                        <td>
                          <span
                            className={`status-badge status-${row.status.toLowerCase()}`}
                          >
                            {PREVIEW_STATUS_TEXT[row.status]}
                          </span>
                        </td>
                        <td>{row.message}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          ) : (
            <p className="placeholder-text">
              选择工资条文件后先执行预览校验，桌面端会本地解析 Excel/CSV
              并标记可发送记录。
            </p>
          )}
        </div>
      </section>

      {sendResult ? (
        <section className="panel-card">
          <div className="panel-title">
            <div>
              <h2>发送结果</h2>
            </div>
          </div>
          <div className="panel-body">
            <div className="summary-grid">
              <div className="summary-pill">总数：{sendResult.totalCount}</div>
              <div className="summary-pill">
                成功：{sendResult.successCount}
              </div>
              <div className="summary-pill">
                失败：{sendResult.failureCount}
              </div>
            </div>
            <div className="table-shell">
              <table>
                <thead>
                  <tr>
                    <th>行号</th>
                    <th>人员</th>
                    <th>邮箱</th>
                    <th>状态</th>
                    <th>说明</th>
                  </tr>
                </thead>
                <tbody>
                  {sendResult.results.map((row) => (
                    <tr key={`${row.rowNumber}-${row.email}-${row.status}`}>
                      <td>{row.rowNumber}</td>
                      <td>{row.recipientName}</td>
                      <td>{row.email}</td>
                      <td>
                        <span
                          className={`status-badge status-${row.status === "SUCCESS" ? "ready" : "invalid"}`}
                        >
                          {SEND_STATUS_TEXT[row.status]}
                        </span>
                      </td>
                      <td>{row.message}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </section>
      ) : null}

      <section className="workspace-grid">
        <article className="panel-card">
          <div className="panel-title">
            <div>
              <h2>邮件模板</h2>
            </div>
          </div>
          <div className="panel-body">
            <label className="field">
              <span>邮件主题模板</span>
              <input
                value={subjectTemplate}
                onChange={(event) => setSubjectTemplate(event.target.value)}
              />
            </label>
            <label className="field">
              <span>邮件正文模板</span>
              <textarea
                rows={12}
                value={bodyTemplate}
                onChange={(event) => setBodyTemplate(event.target.value)}
              />
            </label>
            <p className="helper-text">
              可直接使用表头变量，例如 `{"{{人员}}"}`、`{"{{月份}}"}`、`
              {"{{实发工资}}"}`、`{"{{邮箱}}"}`。
            </p>
            <div className="action-row">
              <button
                className="ghost-button"
                type="button"
                onClick={handleRestoreDefaults}
              >
                恢复默认
              </button>
              <button
                className="ghost-button"
                type="button"
                disabled={isSavingSettings}
                onClick={handleSyncCurrentAsDefault}
              >
                {isSavingSettings ? "保存中..." : "设为默认"}
              </button>
            </div>
          </div>
        </article>

        <article className="panel-card">
          <div className="panel-title">
            <div>
              <h2>SMTP 配置</h2>
            </div>
          </div>
          <div className="panel-body">
            <div className="form-grid">
              <label className="field">
                <span>SMTP 主机</span>
                <input
                  value={smtp.host}
                  onChange={(event) =>
                    setSmtp((current) => ({
                      ...current,
                      host: event.target.value,
                    }))
                  }
                />
              </label>
              <label className="field">
                <span>端口</span>
                <input
                  value={smtp.port}
                  onChange={(event) =>
                    setSmtp((current) => ({
                      ...current,
                      port: event.target.value,
                    }))
                  }
                />
              </label>
              <label className="field">
                <span>用户名</span>
                <input
                  value={smtp.username}
                  onChange={(event) =>
                    setSmtp((current) => ({
                      ...current,
                      username: event.target.value,
                    }))
                  }
                />
              </label>
              <label className="field">
                <span>密码 / 授权码</span>
                <input
                  type="password"
                  value={smtp.password}
                  onChange={(event) =>
                    setSmtp((current) => ({
                      ...current,
                      password: event.target.value,
                    }))
                  }
                />
              </label>
              <label className="field">
                <span>发件邮箱</span>
                <input
                  value={smtp.fromAddress}
                  onChange={(event) =>
                    setSmtp((current) => ({
                      ...current,
                      fromAddress: event.target.value,
                    }))
                  }
                />
              </label>
              <label className="field">
                <span>发件人名称</span>
                <input
                  value={smtp.fromName}
                  onChange={(event) =>
                    setSmtp((current) => ({
                      ...current,
                      fromName: event.target.value,
                    }))
                  }
                />
              </label>
            </div>

            <div className="toggle-row">
              <label className="toggle">
                <input
                  type="checkbox"
                  checked={smtp.auth}
                  onChange={(event) =>
                    setSmtp((current) => ({
                      ...current,
                      auth: event.target.checked,
                    }))
                  }
                />
                <span>启用 SMTP 认证</span>
              </label>
              <label className="toggle">
                <input
                  type="checkbox"
                  checked={smtp.starttls}
                  onChange={(event) =>
                    setSmtp((current) => ({
                      ...current,
                      starttls: event.target.checked,
                    }))
                  }
                />
                <span>启用 STARTTLS</span>
              </label>
            </div>

            <div className="action-row">
              <button
                className="ghost-button"
                type="button"
                onClick={handleRestoreDefaultSmtp}
              >
                恢复默认
              </button>
              <button
                className="ghost-button"
                type="button"
                disabled={isSavingSettings}
                onClick={handleSyncCurrentSmtpAsDefault}
              >
                {isSavingSettings ? "保存中..." : "设为默认"}
              </button>
            </div>
          </div>
        </article>
      </section>
    </main>
  );
}

export default App;
