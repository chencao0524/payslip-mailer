#!/bin/zsh

set -euo pipefail

target="${1:-all}"

build_macos() {
  local arch="$1"
  local bundle="${2:-dmg}"
  ./scripts/build-macos.sh "${arch}" "${bundle}"
}

notarize_macos_if_requested() {
  local arch="$1"
  if [[ "${NOTARIZE_MACOS_DMG:-0}" != "1" ]]; then
    return 0
  fi

  local dmg_path="src-tauri/target/release/bundle/dmg/payslip-mailer-${arch}.dmg"
  if [[ ! -f "${dmg_path}" ]]; then
    echo "Expected DMG not found for notarization: ${dmg_path}" >&2
    exit 1
  fi

  ./scripts/notarize-dmg.sh "${dmg_path}"
}

case "${target}" in
  macos-arm64)
    build_macos arm64 dmg
    notarize_macos_if_requested arm64
    ;;
  macos-x64)
    build_macos x64 dmg
    notarize_macos_if_requested x64
    ;;
  macos)
    build_macos arm64 dmg
    notarize_macos_if_requested arm64
    build_macos x64 dmg
    notarize_macos_if_requested x64
    ;;
  windows)
    ./scripts/build-windows.sh x64 msi
    ;;
  all)
    case "$(uname -s)" in
      Darwin)
        build_macos arm64 dmg
        notarize_macos_if_requested arm64
        build_macos x64 dmg
        notarize_macos_if_requested x64
        echo "Windows packaging is not available on macOS. Run ./scripts/build-release.sh windows on a Windows host." >&2
        ;;
      MINGW*|MSYS*|CYGWIN*|Windows_NT)
        ./scripts/build-windows.sh x64 msi
        ;;
      *)
        echo "Unsupported host OS: $(uname -s)" >&2
        exit 1
        ;;
    esac
    ;;
  *)
    echo "Unsupported target: ${target}" >&2
    echo "Use one of: macos-arm64, macos-x64, macos, windows, all" >&2
    exit 1
    ;;
esac
