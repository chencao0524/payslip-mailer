#!/bin/zsh

set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "Usage: $0 <path-to-dmg>" >&2
  exit 1
fi

dmg_input="${1:A}"
if [[ ! -f "${dmg_input}" ]]; then
  echo "DMG not found: ${dmg_input}" >&2
  exit 1
fi

is_placeholder() {
  local value="${1:-}"
  [[ -z "${value}" || "${value}" == *"你的"* ]]
}

submission_mode=""
if ! is_placeholder "${APPLE_API_KEY:-}" && ! is_placeholder "${APPLE_API_ISSUER:-}" && [[ -n "${APPLE_API_KEY_PATH:-}" ]]; then
  submission_mode="api-key"
elif [[ -n "${APPLE_ID:-}" && -n "${APPLE_PASSWORD:-}" && -n "${APPLE_TEAM_ID:-}" ]]; then
  submission_mode="apple-id"
else
  echo "Missing notarization credentials." >&2
  echo "Provide either APPLE_API_KEY/APPLE_API_ISSUER/APPLE_API_KEY_PATH or APPLE_ID/APPLE_PASSWORD/APPLE_TEAM_ID." >&2
  exit 1
fi

dmg_dir="${dmg_input:h}"
dmg_ext="${dmg_input:e}"
dmg_stem="${dmg_input:t:r}"

# stapler on this machine fails to open the Chinese-named DMG reliably,
# so notarize and staple an ASCII-named copy and keep that as the release artifact.
ascii_stem="$(print -r -- "${dmg_stem}" | iconv -f UTF-8 -t ASCII//TRANSLIT 2>/dev/null || true)"
ascii_stem="$(print -r -- "${ascii_stem}" | tr '[:upper:]' '[:lower:]' | tr -cs 'a-z0-9._-' '-')"
ascii_stem="${ascii_stem#-}"
ascii_stem="${ascii_stem%-}"
if [[ -z "${ascii_stem}" ]]; then
  ascii_stem="payslip-mailer"
fi
ascii_path="${dmg_dir}/${ascii_stem}.${dmg_ext}"

if [[ "${ascii_path}" != "${dmg_input}" ]]; then
  cp -f "${dmg_input}" "${ascii_path}"
fi

echo "Submitting ${ascii_path} for notarization using ${submission_mode} credentials..."
if [[ "${submission_mode}" == "api-key" ]]; then
  xcrun notarytool submit "${ascii_path}" \
    --key "${APPLE_API_KEY_PATH}" \
    --key-id "${APPLE_API_KEY}" \
    --issuer "${APPLE_API_ISSUER}" \
    --wait
else
  xcrun notarytool submit "${ascii_path}" \
    --apple-id "${APPLE_ID}" \
    --password "${APPLE_PASSWORD}" \
    --team-id "${APPLE_TEAM_ID}" \
    --wait
fi

echo "Stapling ticket to ${ascii_path}..."
xcrun stapler staple -v "${ascii_path}"

echo "Validating stapled ticket on ${ascii_path}..."
xcrun stapler validate -v "${ascii_path}"

echo
echo "Notarized DMG ready:"
echo "${ascii_path}"
