# XLS到XLSX批量转换器

一个基于Microsoft Office Interop组件开发的Windows桌面应用程序，用于批量将XLS文件转换为XLSX格式。

## 功能特性

- 🔄 **批量转换**: 支持一次性转换多个XLS文件
- 📁 **文件夹支持**: 可选择包含子文件夹进行递归转换
- 🎯 **智能输出**: 保持原有的文件夹结构
- ⚡ **进度显示**: 实时显示转换进度和状态
- 🛡️ **安全选项**: 可选择是否覆盖已存在的文件
- 🧹 **成功后清理**: 可配置成功转换后自动删除原始XLS文件
- 📡 **自动监控**: 支持后台监听新XLS文件并自动转换
- 🪟 **系统托盘**: 最小化后驻留托盘，双击可快速恢复窗口
- 📝 **详细日志**: 显示每个文件的转换结果并提示状态
- ❌ **取消支持**: 支持中途取消转换操作

## 系统要求

- Windows 7 或更高版本
- Microsoft Excel 2007 或更高版本
- .NET 6.0 或更高版本

## 安装说明

### 方式一：从源码编译

1. 确保已安装 Visual Studio 2022 或 .NET 6.0 SDK
2. 克隆或下载项目源码
3. 打开命令提示符，导航到项目目录
4. 运行以下命令：

```bash
dotnet restore
dotnet build -c Release
```

5. 编译完成后，在 `bin/Release/net6.0-windows/` 目录下找到可执行文件

### 方式二：运行发布版本

```bash
dotnet publish -c Release -r win-x64 --self-contained true
```

## 使用方法

1. **启动程序**: 双击运行 `Xls2XlsxConverter.exe`

2. **选择输入文件夹**: 点击"浏览"按钮选择包含XLS文件的文件夹

3. **选择输出文件夹**: 点击"浏览"按钮选择保存XLSX文件的目标文件夹

4. **设置选项**:
   - ✅ **包含子文件夹**: 勾选后会递归处理所有子文件夹中的XLS文件
   - ✅ **覆盖已存在的文件**: 勾选后会覆盖目标文件夹中已存在的同名XLSX文件
   - ✅ **删除源XLS**: 勾选后转换成功将自动删除原始XLS文件
   - ✅ **自动监控**: 勾选后程序在后台监听输入目录的新XLS并自动转换

5. **开始转换**: 点击"开始转换"按钮启动批量转换

6. **监控进度**: 通过进度条和日志窗口查看转换状态

7. **自动运行** (可选): 最小化到托盘后仍可在后台自动转换新XLS，日志会实时提示

8. **取消操作**: 如需中途停止，点击"取消"按钮

## 项目结构

```
Xls2XlsxConverter/
├── Xls2XlsxConverter.sln          # Visual Studio解决方案文件
├── Xls2XlsxConverter/
│   ├── Xls2XlsxConverter.csproj   # 项目文件
│   ├── Program.cs                 # 程序入口点
│   ├── MainForm.cs                # 主窗体界面
│   └── ExcelConverter.cs          # Excel转换核心逻辑
└── README.md                      # 项目说明文档
```

## 技术实现

### 核心组件

- **MainForm.cs**: Windows Forms主界面，负责用户交互、自动监控、托盘管理等
- **ExcelConverter.cs**: 核心转换逻辑和Excel COM管理，使用 Microsoft.Office.Interop.Excel 进行文件格式转换
- **Program.cs**: 应用程序入口点，负责全局异常处理

### 关键技术

- **Microsoft Office Interop**: 使用Excel COM组件进行文件格式转换
- **异步编程**: 使用async/await模式确保UI响应性
- **COM对象管理**: 正确释放COM对象避免内存泄漏
- **异常处理**: 完善的错误处理和用户提示

### 转换流程

1. 扫描指定文件夹中的所有XLS文件
2. 为每个XLS文件创建对应的XLSX输出路径
3. 使用Excel Interop组件打开XLS文件
4. 将文件保存为XLSX格式
5. 正确释放COM对象资源
6. 更新进度显示和日志记录

## 注意事项

- 转换过程中请勿关闭Excel进程
- 大文件转换可能需要较长时间，请耐心等待
- 确保有足够的磁盘空间存储转换后的文件
- 建议在转换前备份重要文件

## 故障排除

### 常见问题

1. **"无法启动Excel应用程序"错误**
   - 确保已安装Microsoft Excel
   - 检查Excel是否能够正常启动
   - 尝试以管理员身份运行程序

2. **转换失败**
   - 检查源文件是否损坏
   - 确保文件没有被其他程序占用
   - 验证输出文件夹的写入权限

3. **内存占用过高**
   - 避免同时转换过多大文件
   - 关闭不必要的Excel进程

## 开发说明

### 环境配置

```xml
<PackageReference Include="Microsoft.Office.Interop.Excel" Version="15.0.4795.1001" />
```

### 编译命令

```bash
# 调试版本
dotnet build

# 发布版本
dotnet build -c Release

# 创建独立可执行文件
dotnet publish -c Release -r win-x64 --self-contained true
```

## 版本历史

- **v1.1.0** (2025-10-21)
  - 新增自动监控输入目录并自动转换的功能
  - 支持转换成功后自动删除源XLS文件（可选）
  - 窗口最小化后驻留系统托盘，双击可快速恢复
  - 新增删除目录相同的校验与日志增强

- **v1.0.0** (2025-10-21)
  - 首次发布
  - 支持批量XLS到XLSX转换
  - 图形化用户界面
  - 进度显示和日志记录

## 许可证

本项目采用 MIT 许可证。详情请参阅 LICENSE 文件。

## 技术支持

如有问题或建议，请联系开发团队。