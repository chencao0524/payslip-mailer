#!/bin/zsh

set -euo pipefail

arch="${1:-arm64}"
bundle="${2:-dmg}"

case "${arch}" in
  arm64)
    target="aarch64-apple-darwin"
    ;;
  x64)
    target="x86_64-apple-darwin"
    ;;
  *)
    echo "Unsupported architecture: ${arch}" >&2
    echo "Use one of: arm64, x64" >&2
    exit 1
    ;;
esac

case "${bundle}" in
  app|dmg)
    ;;
  *)
    echo "Unsupported bundle type: ${bundle}" >&2
    echo "Use one of: app, dmg" >&2
    exit 1
    ;;
esac

echo "Building macOS ${bundle} for ${arch} (${target})..."

if [[ -n "${APPLE_CERTIFICATE_FILE:-}" && -z "${APPLE_CERTIFICATE:-}" ]]; then
  if [[ ! -f "${APPLE_CERTIFICATE_FILE}" ]]; then
    echo "APPLE_CERTIFICATE_FILE not found: ${APPLE_CERTIFICATE_FILE}" >&2
    exit 1
  fi
  export APPLE_CERTIFICATE="$(base64 -i "${APPLE_CERTIFICATE_FILE}")"
fi

export TAURI_BUNDLE_TARGET="${target}"
node node_modules/@tauri-apps/cli/tauri.js build --target "${target}" --bundles "${bundle}"

if [[ "${bundle}" == "dmg" ]]; then
  host_target="$(rustc -vV | awk '/host:/ { print $2 }')"
  source_dmg_dir="src-tauri/target/release/bundle/dmg"
  if [[ "${target}" != "${host_target}" && -d "src-tauri/target/${target}/release/bundle/dmg" ]]; then
    source_dmg_dir="src-tauri/target/${target}/release/bundle/dmg"
  fi

  latest_dmg="$(ls -t "${source_dmg_dir}"/*.dmg 2>/dev/null | head -n 1 || true)"
  if [[ -z "${latest_dmg}" ]]; then
    echo "Could not find a generated DMG in ${source_dmg_dir}" >&2
    exit 1
  fi

  output_dmg_dir="src-tauri/target/release/bundle/dmg"
  mkdir -p "${output_dmg_dir}"
  ascii_dmg="${output_dmg_dir}/payslip-mailer-${arch}.dmg"
  if [[ "${latest_dmg}" != "${ascii_dmg}" ]]; then
    cp -f "${latest_dmg}" "${ascii_dmg}"
    echo "Created ASCII DMG copy for notarization compatibility:"
    echo "${ascii_dmg}"
  fi
fi
