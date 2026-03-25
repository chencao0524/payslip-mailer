import { invoke } from "@tauri-apps/api/core";
import {
  buildDefaultBodyTemplate,
  type ContactMappingStore,
  DEFAULT_SMTP,
  DEFAULT_SUBJECT_TEMPLATE,
  type PayslipSettings,
  type SendRequest,
  type SendResponse,
} from "./payslip";

const SETTINGS_KEY = "payslip-mailer-settings";
const CONTACT_MAPPING_KEY = "payslip-mailer-contact-mapping";

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

export async function loadContactMapping() {
  if (isTauriRuntime()) {
    return invoke<ContactMappingStore | null>("payslip_get_contact_mapping");
  }
  const raw = window.localStorage.getItem(CONTACT_MAPPING_KEY);
  if (!raw) {
    return null;
  }
  try {
    return JSON.parse(raw) as ContactMappingStore;
  } catch {
    return null;
  }
}

export async function saveContactMapping(mapping: ContactMappingStore) {
  if (isTauriRuntime()) {
    return invoke<ContactMappingStore>("payslip_save_contact_mapping", {
      mapping,
    });
  }
  window.localStorage.setItem(CONTACT_MAPPING_KEY, JSON.stringify(mapping));
  return mapping;
}

export async function clearContactMapping() {
  if (isTauriRuntime()) {
    return invoke<void>("payslip_clear_contact_mapping");
  }
  window.localStorage.removeItem(CONTACT_MAPPING_KEY);
}

export async function readLocalFile(path: string) {
  return invoke<{ fileName: string; bytes: number[] }>("payslip_read_local_file", {
    path,
  });
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
