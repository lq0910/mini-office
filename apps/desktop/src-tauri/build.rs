fn main() {
    // 功能：前端构建变化时重新嵌入资源，避免只重打包旧二进制；时间：2026-11-14；作者：lq。
    println!("cargo:rerun-if-changed=../dist");
    tauri_build::build()
}
