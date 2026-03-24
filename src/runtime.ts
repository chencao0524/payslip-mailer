import { invoke } from "@tauri-apps/api/core";
import {
  buildDefaultBodyTemplate,
  DEFAULT_SMTP,
  DEFAULT_SUBJECT_TEMPLATE,
  type PayslipSettings,
  type SendRequest,
  type SendResponse,
} from "./payslip";

const SETTINGS_KEY = "payslip-mailer-settings";

export function isTauriRuntime() {
  if (typeof window === "undefined") {
    return false;
  }
  return "__TAURI_INTERNALS__" in window || navigator.userAgent.includes("Tauri");
}

function fallbackSettings(): PayslipSettings {
  return {
    defaultSubjectTemplate: DEFAULT_SUBJECT_TEMPLATE,
    defaultBodyTemplate: buildDefaultBodyTemplate(),
    smtp: DEFAULT_SMTP,
  };
}

export async function loadSettings() {
  if (isTauriRuntime()) {
    return invoke<PayslipSettings>("payslip_get_settings");
  }
  const raw = window.localStorage.getItem(SETTINGS_KEY);
  if (!raw) {
    return fallbackSettings();
  }
  try {
    return JSON.parse(raw) as PayslipSettings;
  } catch {
    return fallbackSettings();
  }
}

export async function saveSettings(settings: PayslipSettings) {
  if (isTauriRuntime()) {
    return invoke<PayslipSettings>("payslip_save_settings", { settings });
  }
  window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  return settings;
}

export async function sendPayslips(request: SendRequest) {
  if (isTauriRuntime()) {
    return invoke<SendResponse>("payslip_send", { request });
  }
  await new Promise((resolve) => window.setTimeout(resolve, 320));
  return {
    totalCount: request.rows.length,
    successCount: request.rows.length,
    failureCount: 0,
    results: request.rows.map((row) => ({
      rowNumber: row.rowNumber,
      recipientName: row.recipientName,
      email: row.email,
      status: "SUCCESS" as const,
      message: "Web 调试模式未实际发信，已模拟成功。",
      values: row.values,
    })),
  } satisfies SendResponse;
}
