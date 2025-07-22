# 供应商对帐工具集

## 项目说明
这是一个用于供应商对账的工具集，包含以下功能：
- 对账明细处理
- 商品分类管理

## 系统要求
- Windows 操作系统
- Python 3.11 或更高版本（如果从源码运行）

## 依赖项
```
pandas>=2.0.0
openpyxl>=3.1.0
ttkbootstrap>=1.10.1
pyinstaller>=6.0.0
```

## 安装和使用
1. 从 Release 页面下载最新版本的可执行文件
2. 解压缩下载的文件
3. 运行 `供应商对帐工具集.exe`

## 开发说明
如果要从源码运行或开发：
1. 克隆仓库：
```bash
git clone https://github.com/[your-username]/New_Recon_Tools.git
```
2. 安装依赖：
```bash
pip install -r requirements.txt
```
3. 运行主程序：
```bash
python Bldbuy_Recon_UI.py
```

## 构建
使用以下命令构建可执行文件：
```bash
python build_with_version.py
```

## 版本历史
- v2.0.0: 初始发布版本
  - 实现基本的对账功能
  - 添加商品分类管理