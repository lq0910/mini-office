#!/usr/bin/env bash
# 功能：本机生成签名应用、标准 DMG、ZIP 和校验清单，避免依赖 Finder AppleScript；时间：2026-11-11；作者：lq。
set -euo pipefail
cd "$(dirname "$0")/.."
if [[ "${MINI_OFFICE_SKIP_NATIVE_BUILD:-0}" != "1" ]]; then
  APPLE_SIGNING_IDENTITY="${APPLE_SIGNING_IDENTITY:--}" pnpm tauri:build --bundles app --ci
fi
app_path="target/release/bundle/macos/Mini Office.app"
# 功能：拒绝源码已变化的旧前端，并确认新资源确实进入原生二进制；时间：2026-11-14；作者：lq。
build_id=$(node scripts/build-fingerprint.mjs --check)
python3 - "$app_path/Contents/MacOS/mini-office-desktop" <<'PY'
# 功能：核对入口与 Word 编辑器资源名，防止 dist 新而可执行文件仍旧；时间：2026-11-14；作者：lq。
import hashlib, json, pathlib, sys
info = json.loads(pathlib.Path('apps/desktop/dist/build-info.json').read_text())
binary = pathlib.Path(sys.argv[1]).read_bytes()
# 功能：打包前校验所有最终 Web 资源哈希，任何旧文件或构建后改写都立即失败；时间：2026-11-14；作者：lq。
for name, expected in info['assets'].items():
    actual = hashlib.sha256((pathlib.Path('apps/desktop/dist') / name).read_bytes()).hexdigest()
    if actual != expected:
        raise SystemExit(f'前端产物与最终构建清单不一致：{name}')
    if name.startswith(('assets/index-', 'assets/WordCoreEditor-')) and name.endswith(('.js', '.css')):
        if name.encode() not in binary:
            raise SystemExit(f'原生二进制缺少本次前端资源：{name}')
PY
release_version=$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "$app_path/Contents/Info.plist")
release_arch=$(lipo -archs "$app_path/Contents/MacOS/mini-office-desktop")
case "$release_arch" in
  arm64) release_suffix=arm64 ;;
  x86_64) release_suffix=x64 ;;
  'x86_64 arm64'|'arm64 x86_64') release_suffix=universal ;;
  *) echo "不支持的架构：$release_arch" >&2; exit 1 ;;
esac
release_root="$PWD/output/releases/$release_version"
mkdir -p "$release_root"
staging_root=$(mktemp -d "$release_root/.dmg-root.XXXXXX")
# 功能：只清理由本脚本创建的临时镜像目录；时间：2026-11-11；作者：lq。
trap 'rm -rf "$staging_root"' EXIT
codesign --verify --deep --strict "$app_path"
ditto "$app_path" "$staging_root/Mini Office.app"
ln -s /Applications "$staging_root/Applications"
release_name="Mini-Office_${release_version}_macos-${release_suffix}_${build_id}"
# 功能：构建信息随包交付，包名包含唯一编号，避免同名旧包混淆；时间：2026-11-14；作者：lq。
cp apps/desktop/dist/build-info.json "$release_root/$release_name.build.json"
hdiutil create -ov -volname "Mini Office $release_version" -srcfolder "$staging_root" -format UDZO -o "$release_root/$release_name.dmg"
ditto -c -k --sequesterRsrc --keepParent "$app_path" "$release_root/$release_name.zip"
hdiutil verify "$release_root/$release_name.dmg"
cd "$release_root"
shasum -a 256 "$release_name.dmg" "$release_name.zip" "$release_name.build.json" > "$release_name.SHA256SUMS.txt"
shasum -a 256 -c "$release_name.SHA256SUMS.txt"
