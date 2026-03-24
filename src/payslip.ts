import * as XLSX from "xlsx";

export type PreviewStatus = "READY" | "INVALID" | "SKIPPED";

export type PreviewRow = {
  rowNumber: number;
  recipientName: string;
  email: string;
  month: string;
  netPay: string;
  status: PreviewStatus;
  message: string;
  values: Record<string, string>;
};

export type PreviewResponse = {
  fileName: string;
  sheetName: string;
  headers: string[];
  rows: PreviewRow[];
  totalCount: number;
  readyCount: number;
  invalidCount: number;
  skippedCount: number;
  duplicateEmailMessages: string[];
  defaultSubjectTemplate: string;
  defaultBodyTemplate: string;
};

export type SendResult = {
  rowNumber: number;
  recipientName: string;
  email: string;
  status: "SUCCESS" | "FAILED";
  message: string;
  values: Record<string, string>;
};

export type SendResponse = {
  totalCount: number;
  successCount: number;
  failureCount: number;
  results: SendResult[];
};

export type SendRequest = {
  subjectTemplate: string;
  bodyTemplate: string;
  smtp: SmtpSettings;
  rows: Array<{
    rowNumber: number;
    recipientName: string;
    email: string;
    values: Record<string, string>;
  }>;
};

export type SmtpSettings = {
  host: string;
  port: string;
  username: string;
  password: string;
  fromAddress: string;
  fromName: string;
  auth: boolean;
  starttls: boolean;
};

export type PayslipSettings = {
  defaultSubjectTemplate: string;
  defaultBodyTemplate: string;
  smtp: SmtpSettings;
};

const EMAIL_HEADER = "邮箱";
const NAME_HEADER = "人员";
const MONTH_HEADER = "月份";
const NET_PAY_HEADER = "实发工资";
const GROSS_PAY_HEADER = "工资总额";
const REQUIRED_HINT_HEADERS = [EMAIL_HEADER, NAME_HEADER, MONTH_HEADER];
const EMAIL_PATTERN = /^[A-Za-z0-9+_.-]+@[A-Za-z0-9.-]+$/;

export const DEFAULT_SUBJECT_TEMPLATE = "【薪酬通知】{{月份}}月工资条 - {{人员}}";

export const DEFAULT_TEMPLATE_HEADERS = [
  "备注",
  "序号",
  "月份",
  "公司",
  "人员",
  "应 出勤",
  "实际 出勤",
  "二级部门",
  "三级部门",
  "固定基本工资",
  "岗位工资",
  "绩效奖金",
  "补贴",
  "新增/试用 调整",
  "离职调整",
  "电脑补贴",
  "特殊津贴 月度提成",
  "派驻补贴",
  "差旅补贴",
  "离职补偿",
  "兼职津贴",
  "奖金",
  "加班津贴",
  "事假天数",
  "病假扣款",
  "事假扣款",
  "迟到早退 未打卡扣款",
  "其他扣款",
  "补发 或 扣发",
  "工资总额",
  "养老保险",
  "医疗保险",
  "失业保险",
  "公积金",
  "个人小计",
  "本期累计预扣预缴应税额",
  "适用税率",
  "速算 扣除数",
  "本期累计代扣代缴 个人所得税",
  "往期累计已预扣预缴税额",
  "往来扣款",
  "其他",
  "实发工资",
  "本期实际应预扣预缴税额税额",
] as const;

export const DEFAULT_SMTP: SmtpSettings = {
  host: "smtp.mxhichina.com",
  port: "465",
  username: "",
  password: "",
  fromAddress: "",
  fromName: "工资条通知",
  auth: true,
  starttls: false,
};

export function buildDefaultBodyTemplate() {
  const rows = DEFAULT_TEMPLATE_HEADERS.map(
    (header) =>
      `  <tr><td style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">${header}</td><td style="padding: 6px 8px; border: 1px solid #d9dee8;">{{${header}}}</td></tr>`,
  ).join("\n");
  return `
<p style="margin: 0 0 12px;">{{人员}}，您好：</p>

<p style="margin: 0 0 12px;">现将您 <strong>{{月份}} 月</strong> 工资发放情况通知如下，请您查收并妥善保管。</p>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; width: auto; max-width: 100%; border-color: #d9dee8; font-size: 12px; line-height: 1.45; margin: 0 0 12px;">
  <tr style="background: #f7f9fc;">
    <th align="left" width="132" style="padding: 6px 8px; border: 1px solid #d9dee8; white-space: nowrap;">项目</th>
    <th align="left" style="padding: 6px 8px; border: 1px solid #d9dee8;">内容</th>
  </tr>
${rows}
</table>

<p style="margin: 0 0 12px;">如您对本人工资条信息有疑问，请在收到邮件后及时与财务或人力资源相关同事联系核实。</p>
<p style="margin: 0 0 12px;">感谢您的配合。</p>
<p style="margin: 0 0 12px;">此致<br/>敬礼</p>
<p style="margin: 0 0 12px;">财务部 / 人力资源部<br/>工资条通知系统</p>
<p style="margin: 0; color: #6b7280; font-size: 12px;">本邮件包含个人薪酬敏感信息，仅限收件人本人查阅，请勿擅自转发、传播或用于其他用途。</p>
`.trim();
}

export async function parsePayslipFile(file: File): Promise<PreviewResponse> {
  const lowerName = file.name.toLowerCase();
  if (lowerName.endsWith(".csv")) {
    return parseCsvFile(file);
  }
  if (lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls")) {
    return parseWorkbookFile(file);
  }
  throw new Error("仅支持 xlsx、xls 或 csv 文件");
}

function parseWorkbookFile(file: File): Promise<PreviewResponse> {
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: "array", raw: false, cellText: false, cellDates: false });
    const sheetName = workbook.SheetNames.includes("员工工资条") ? "员工工资条" : workbook.SheetNames[0];
    if (!sheetName) {
      throw new Error("工资条工作表不存在");
    }

    const rows = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(workbook.Sheets[sheetName], {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false,
    });

    const headerRowIndex = locateHeaderRow(rows);
    if (headerRowIndex < 0) {
      throw new Error("未识别到工资条表头，请确认包含“人员、月份、邮箱”等列");
    }

    return buildPreviewResponse(file.name, sheetName, rows, headerRowIndex);
  });
}

async function parseCsvFile(file: File): Promise<PreviewResponse> {
  const content = await file.text();
  const workbook = XLSX.read(content, { type: "string", raw: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("CSV 文件为空");
  }
  const rows = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(workbook.Sheets[sheetName], {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  });
  const headerRowIndex = locateHeaderRow(rows);
  if (headerRowIndex < 0) {
    throw new Error("未识别到工资条表头，请确认包含“人员、月份、邮箱”等列");
  }
  return buildPreviewResponse(file.name, "CSV", rows, headerRowIndex);
}

function buildPreviewResponse(fileName: string, sheetName: string, rawRows: (string | number | boolean | null)[][], headerRowIndex: number) {
  const headers = trimTrailingBlankHeaders(
    (rawRows[headerRowIndex] ?? []).map((value) => normalizeText(String(value ?? ""))),
  );

  const rows: PreviewRow[] = [];
  for (let rowIndex = headerRowIndex + 1; rowIndex < rawRows.length; rowIndex += 1) {
    const row = rawRows[rowIndex] ?? [];
    const values: Record<string, string> = {};
    let hasAnyValue = false;

    for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
      const header = headers[columnIndex];
      if (!header) {
        continue;
      }
      const value = normalizeText(String(row[columnIndex] ?? ""));
      values[header] = value;
      hasAnyValue = hasAnyValue || value.length > 0;
    }

    if (!hasAnyValue) {
      continue;
    }

    rows.push(toPreviewRow(rowIndex + 1, values));
  }

  const duplicateEmailMessages = markDuplicateEmails(rows);

  const readyCount = rows.filter((row) => row.status === "READY").length;
  const invalidCount = rows.filter((row) => row.status === "INVALID").length;
  const skippedCount = rows.filter((row) => row.status === "SKIPPED").length;

  return {
    fileName,
    sheetName,
    headers,
    rows,
    totalCount: rows.length,
    readyCount,
    invalidCount,
    skippedCount,
    duplicateEmailMessages,
    defaultSubjectTemplate: DEFAULT_SUBJECT_TEMPLATE,
    defaultBodyTemplate: buildDefaultBodyTemplate(),
  } satisfies PreviewResponse;
}

function locateHeaderRow(rows: (string | number | boolean | null)[][]) {
  const limit = Math.min(rows.length, 12);
  let bestIndex = -1;
  let bestScore = -1;
  for (let index = 0; index < limit; index += 1) {
    const values = (rows[index] ?? []).map((value) => normalizeText(String(value ?? "")));
    const score = scoreHeaderRow(values);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  }
  return bestScore >= 10 ? bestIndex : -1;
}

function scoreHeaderRow(values: string[]) {
  const nonBlankCount = values.filter(Boolean).length;
  let score = Math.min(nonBlankCount, 12);
  REQUIRED_HINT_HEADERS.forEach((required) => {
    if (values.includes(required)) {
      score += 8;
    }
  });
  if (values.includes(NET_PAY_HEADER) || values.includes(GROSS_PAY_HEADER)) {
    score += 6;
  }
  return score;
}

function trimTrailingBlankHeaders(headers: string[]) {
  let lastNonBlank = headers.length - 1;
  while (lastNonBlank >= 0 && !headers[lastNonBlank]) {
    lastNonBlank -= 1;
  }
  return headers.slice(0, lastNonBlank + 1);
}

function toPreviewRow(rowNumber: number, values: Record<string, string>): PreviewRow {
  const recipientName = valueOf(values, NAME_HEADER);
  const month = valueOf(values, MONTH_HEADER);
  const email = valueOf(values, EMAIL_HEADER);
  const netPay = valueOf(values, NET_PAY_HEADER);
  const grossPay = valueOf(values, GROSS_PAY_HEADER);

  if (!recipientName || !month || (!netPay && !grossPay)) {
    return { rowNumber, recipientName, email, month, netPay, status: "SKIPPED", message: "非工资明细行，已跳过", values };
  }
  if (!email) {
    return { rowNumber, recipientName, email, month, netPay, status: "INVALID", message: "邮箱为空，无法发送", values };
  }
  if (!EMAIL_PATTERN.test(email)) {
    return { rowNumber, recipientName, email, month, netPay, status: "INVALID", message: "邮箱格式不正确", values };
  }
  return { rowNumber, recipientName, email, month, netPay, status: "READY", message: "校验通过，可发送", values };
}

function markDuplicateEmails(rows: PreviewRow[]) {
  const emailMap = new Map<string, PreviewRow[]>();

  rows.forEach((row) => {
    const normalizedEmail = normalizeEmail(row.email);
    if (!normalizedEmail) {
      return;
    }
    const group = emailMap.get(normalizedEmail);
    if (group) {
      group.push(row);
      return;
    }
    emailMap.set(normalizedEmail, [row]);
  });

  const duplicates = Array.from(emailMap.entries())
    .filter(([, group]) => group.length > 1)
    .sort(([left], [right]) => left.localeCompare(right, "zh-CN"));

  duplicates.forEach(([, group]) => {
    group.forEach((row) => {
      row.status = "INVALID";
      row.message = "邮箱重复，已禁止发送，请先修正导入文件";
    });
  });

  return duplicates.map(([email, group]) => {
    const rowNumbers = group.map((row) => row.rowNumber).join("、");
    const names = group
      .map((row) => row.recipientName)
      .filter(Boolean)
      .join(" / ");
    return names
      ? `${email}（第 ${rowNumbers} 行，${names}）`
      : `${email}（第 ${rowNumbers} 行）`;
  });
}

function valueOf(values: Record<string, string>, key: string) {
  return normalizeText(values[key] ?? "");
}

function normalizeEmail(value: string) {
  return normalizeText(value).toLowerCase();
}

function normalizeText(value: string) {
  return value
    .replace(/\uFEFF/g, "")
    .replace(/\n/g, " ")
    .replace(/\r/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function escapeCsvCell(value: string) {
  const normalized = (value ?? "").replace(/"/g, "\"\"");
  return /[",\n]/.test(normalized) ? `"${normalized}"` : normalized;
}

export function exportFailedRows(preview: PreviewResponse, failedRows: Array<{
  rowNumber: number;
  recipientName: string;
  email: string;
  message: string;
  values: Record<string, string>;
}>) {
  if (failedRows.length === 0) {
    return;
  }
  const headers = ["失败原因", ...preview.headers.filter(Boolean)];
  const csvRows = [
    headers,
    ...failedRows.map((row) => [
      row.message,
      ...preview.headers.filter(Boolean).map((header) => row.values[header] ?? ""),
    ]),
  ];
  const content = csvRows.map((row) => row.map((cell) => escapeCsvCell(cell ?? "")).join(",")).join("\n");
  const blob = new Blob(["\uFEFF" + content], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${preview.fileName.replace(/\.[^.]+$/, "") || "工资条失败记录"}-失败记录.csv`;
  link.click();
  URL.revokeObjectURL(url);
}
