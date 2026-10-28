# -*- coding: utf-8 -*-
"""打包脚本：在 PyInstaller 编译完成后运行，生成可分发安装包 zip。"""
import os, sys, shutil, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(HERE, "dist", "GaoKaoTrainer")
EXE  = os.path.join(DIST, "GaoKaoTrainer.exe")
RELEASE_NAME = "高考学习训练系统_安装包"
RELEASE_DIR = os.path.join(HERE, RELEASE_NAME)

def fail(msg):
    print("[FAIL]", msg); sys.exit(1)

def main():
    if not os.path.exists(EXE):
        fail("未找到编译产物: %s\n请先运行: pyinstaller --name GaoKaoTrainer ..." % EXE)
    print("[OK] 编译产物已就绪:", EXE)

    # 1. 清理旧包
    if os.path.exists(RELEASE_DIR):
        shutil.rmtree(RELEASE_DIR)
    os.makedirs(os.path.join(RELEASE_DIR, "dist", "GaoKaoTrainer"))

    # 2. 复制编译产物
    shutil.copytree(DIST, os.path.join(RELEASE_DIR, "dist", "GaoKaoTrainer"),
                    dirs_exist_ok=True)
    print("[OK] 编译产物已复制到安装包目录")

    # 3. 复制安装/卸载脚本
    for fname in ["安装.bat", "卸载.bat"]:
        src = os.path.join(HERE, "installer", fname)
        if os.path.exists(src):
            shutil.copy(src, os.path.join(RELEASE_DIR, fname))
    # 卸载脚本也放到 dist 内供安装时复制
    uninstall_src = os.path.join(HERE, "installer", "卸载.bat")
    if os.path.exists(uninstall_src):
        shutil.copy(uninstall_src, os.path.join(RELEASE_DIR, "dist", "GaoKaoTrainer", "卸载.bat"))
    print("[OK] 安装/卸载脚本已复制")

    # 4. 生成使用说明（简版）
    readme = os.path.join(RELEASE_DIR, "使用说明.txt")
    with open(readme, "w", encoding="utf-8") as f:
        f.write("高考学习训练系统（简化版）安装包\n")
        f.write("================================\n\n")
        f.write("安装：双击 安装.bat → 自动创建开始菜单和桌面快捷方式\n")
        f.write("卸载：运行开始菜单中 高考学习训练系统 → 卸载.bat\n")
        f.write("       （用户数据如需保留，卸载时不会被删除）\n\n")
        f.write("默认账号：teacher / 123456（教师）  student / 123456（学生）\n")
        f.write("数据目录：%LOCALAPPDATA%\\GaoKaoTrainer\\data\\\n")
        f.write("大模型API：启动后 → 教师端 → API设置 → 填入 base_url / API Key / 模型名\n")
    print("[OK] 使用说明已生成")

    # 5. 打包为 ZIP
    zip_path = os.path.join(HERE, RELEASE_NAME + ".zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(RELEASE_DIR):
            for fn in files:
                abspath = os.path.join(root, fn)
                arcname = os.path.relpath(abspath, HERE)
                zf.write(abspath, arcname)
    print("[OK] 安装包已生成:", zip_path)
    print("[OK] 安装目录:", RELEASE_DIR)
    print("\n=== 交付物 ===")
    print("  ZIP 文件:", zip_path)
    print("  使用说明: 解压后双击 安装.bat 即可。")

if __name__ == "__main__":
    main()
