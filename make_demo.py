# -*- coding: utf-8 -*-
"""创建演示项目数据:项目文件夹 + BOM Excel + 物料资料"""
import os

PROJ_DIR = r"D:\物料清单\HY2026062301-浙江海港智慧能源有限公司数智跨境充电桩项目铅酸电池充电仓项目"
MAT_DIR = os.path.join(PROJ_DIR, "物料资料")
os.makedirs(os.path.join(MAT_DIR, "8S131100AA-总进线智能电表", "图纸"), exist_ok=True)
os.makedirs(os.path.join(MAT_DIR, "8S131100AA-总进线智能电表", "技术文档"), exist_ok=True)
os.makedirs(os.path.join(MAT_DIR, "8S131200BB"), exist_ok=True)

# 1. BOM Excel
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "总BOM清单"
ws.append(["物料号", "名称", "规格型号", "数量", "单位"])
ws.append(["8S131100AA", "总进线智能电表", "DTSD341-MC3 三相四线", "1", "台"])
ws.append(["8S131200BB", "铅酸电池充电仓", "HC-48V/100A", "2", "套"])
ws.append(["8S131300CC", "直流接触器", "CZ0-400/20", "6", "只"])
ws.append(["8S131400DD", "智能断路器", "CW1-2000/3P", "3", "台"])
ws.append(["8S131500EE", "充电控制模块", "CM-220V", "12", "块"])
bom_path = os.path.join(PROJ_DIR, "总BOM清单.xlsx")
wb.save(bom_path)
print("BOM 已保存:", bom_path)

# 2. 模拟已有资料
files = {
    os.path.join(MAT_DIR, "8S131100AA-总进线智能电表", "图纸", "安装尺寸图.pdf"): "安装尺寸图(演示数据)",
    os.path.join(MAT_DIR, "8S131100AA-总进线智能电表", "技术文档", "使用说明书.docx"): "使用说明书(演示数据)",
    os.path.join(MAT_DIR, "8S131200BB", "充电仓技术协议.pdf"): "技术协议(演示数据)",
}
for p, content in files.items():
    with open(p, "w", encoding="utf-8") as f:
        f.write(content)
    print("资料:", os.path.relpath(p, PROJ_DIR))

# 3. 把用户已有的散资料文件夹也纳入项目物料资料(不动原文件,复制参考)
src = r"D:\物料清单\8S131100AA-总进线智能电表"
dst = os.path.join(MAT_DIR, "8S131100AA-总进线智能电表", "历史资料")
if os.path.isdir(src):
    import shutil
    os.makedirs(dst, exist_ok=True)
    for fn in os.listdir(src):
        s = os.path.join(src, fn)
        if os.path.isfile(s):
            shutil.copy2(s, os.path.join(dst, fn))
            print("已并入历史资料:", fn)

print("\n目录结构:")
for root, dirs, files in os.walk(PROJ_DIR):
    level = root.replace(PROJ_DIR, "").count(os.sep)
    print("  " * level + os.path.basename(root) + "/")
    for fn in files:
        print("  " * (level + 1) + fn)
