# -*- mode: python ; coding: utf-8 -*-

import os
import tempfile
from PIL import Image

def log(message):
    """打印带时间戳的日志"""
    from datetime import datetime
    timestamp = datetime.now().strftime('%H:%M:%S')
    print(f"[{timestamp}] {message}")

# 生成新的图标
png_path = 'icon.png'
ico_path = os.path.join(tempfile.gettempdir(), 'build_icon.ico')
log(f"源图标文件: {png_path}")
log(f"目标ICO文件: {ico_path}")

if os.path.exists(png_path):
    log(f"找到源图标文件: {png_path}")
    img = Image.open(png_path)
    log(f"图标尺寸: {img.size}, 模式: {img.mode}")
    
    if img.mode != 'RGBA':
        log("转换图像模式为 RGBA")
        img = img.convert('RGBA')
    
    # 移除白色背景，使用更精确的阈值
    log("开始处理白色背景...")
    data = img.getdata()
    new_data = []
    white_pixels = 0
    for item in data:
        if item[0] > 250 and item[1] > 250 and item[2] > 250:
            new_data.append((255, 255, 255, 0))
            white_pixels += 1
        else:
            new_data.append(item)
    img.putdata(new_data)
    log(f"已 {white_pixels} 个白色像素")
    
    # 优化中等尺寸的图标处理
    icon_sizes = [
        (1024, 1024),  # 原始尺寸
        (512, 512),    # 超高分辨率
        (256, 256),    # 高分辨率
        (192, 192),    # 新增尺寸
        (128, 128),    # 大尺寸
        (96, 96),      # 中等尺寸
        (72, 72),      # 新增尺寸
        (64, 64),      # 显示尺寸
        (48, 48),      # 显示尺寸
        (40, 40),      # Windows 任务栏
        (32, 32),      # 标准尺寸
        (24, 24),      # Windows 任务栏
        (20, 20),      # Windows 任务栏
        (16, 16)       # 最小尺寸
    ]
    log(f"准备生成 {len(icon_sizes)} 种尺寸的图标")
    
    icons = []
    for size in icon_sizes:
        log(f"处理 {size[0]}x{size[1]} 尺寸...")
        icon = Image.new('RGBA', size, (0, 0, 0, 0))
        
        # 使用两步缩放来提高质量
        if size[0] < img.width:
            # 第一步：先缩放到目标尺寸的2倍
            intermediate_size = (size[0] * 2, size[1] * 2)
            temp_img = img.resize(intermediate_size, Image.Resampling.LANCZOS)
            
            # 第二步：缩放到最终尺寸
            temp_img = temp_img.resize(size, Image.Resampling.LANCZOS)
        else:
            temp_img = img.resize(size, Image.Resampling.LANCZOS)
        
        # 计算居中位置
        x = (size[0] - temp_img.width) // 2
        y = (size[1] - temp_img.height) // 2
        
        # 使用 alpha 通道作为蒙版
        icon.paste(temp_img, (x, y), temp_img)
        
        # 对不同尺寸使用不同的图像增强
        if 32 <= size[0] <= 128:
            from PIL import ImageEnhance
            # 增强对比度
            contrast = ImageEnhance.Contrast(icon)
            icon = contrast.enhance(1.2)
            # 增强锐度
            sharpness = ImageEnhance.Sharpness(icon)
            icon = sharpness.enhance(1.5)
            # 增强颜色
            color = ImageEnhance.Color(icon)
            icon = color.enhance(1.1)
            log(f"已对 {size[0]}x{size[1]} 尺寸应用增强处理")
            
        icons.append(icon)
    
    # 修改保存部分的代码
    log("开始保存 ICO 文件...")
    
    # 将图标分组，每组最多4个尺寸
    def chunk_sizes(lst, n):
        for i in range(0, len(lst), n):
            yield lst[i:i + n]
    
    # 创建临时 ICO 文件
    temp_ico_files = []
    for i, size_group in enumerate(chunk_sizes(icons, 4)):
        temp_path = os.path.join(tempfile.gettempdir(), f'temp_icon_{i}.ico')
        temp_ico_files.append(temp_path)
        
        sizes_in_group = [(icon.width, icon.height) for icon in size_group]
        log(f"保存临时图标组 {i + 1}, 尺寸: {sizes_in_group}")
        
        try:
            size_group[0].save(
                temp_path,
                format='ICO',
                sizes=sizes_in_group,
                append_images=size_group[1:],
                optimize=True,
                quality=100
            )
        except Exception as e:
            log(f"保存临时图标组 {i + 1} 时出错: {str(e)}")
    
    # 合并所有临时 ICO 文件
    try:
        # 读取所有临时文件中的图标
        all_icons = []
        for temp_file in temp_ico_files:
            with Image.open(temp_file) as ico:
                for size in ico.info.get('sizes', []):
                    ico.size = size
                    all_icons.append(ico.copy())
        
        # 保存最终的 ICO 文件
        if all_icons:
            log(f"合并所有图标到最终文件: {ico_path}")
            all_icons[0].save(
                ico_path,
                format='ICO',
                sizes=[(icon.width, icon.height) for icon in all_icons],
                append_images=all_icons[1:],
                optimize=True,
                quality=100
            )
            
            # 验证最文
            with Image.open(ico_path) as final_ico:
                available_sizes = final_ico.info.get('sizes', [])
                log(f"最终 ICO 文件包含的尺寸: {available_sizes}")
                log(f"ICO 文件大小: {os.path.getsize(ico_path)} 字节")
        else:
            log("错误: 没有可用的图标")
            
    except Exception as e:
        log(f"合并图标文件时出错: {str(e)}")
    finally:
        # 清理临时文件
        for temp_file in temp_ico_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except:
                pass
else:
    log(f"警告: 找不到源图标文件 {png_path}")

# 确保 ICO 文件存在
if not os.path.exists(ico_path):
    log(f"错误: ICO 文件不存在 {ico_path}")
    raise FileNotFoundError(f"Icon file {ico_path} not found!")
else:
    log(f"确认 ICO 文件存在: {ico_path}")

# 定义加密块
block_cipher = None

# poppler 相关配置
poppler_path = 'D:\\Program Files\\poppler-24.08.0\\Library\\bin'

# 只收集实际存在的核心 DLL 文件
core_dlls = [
    'poppler-glib.dll',
    'poppler-cpp.dll',
    'poppler.dll',
    'freetype.dll',
    'libpng16.dll',
    'libtiff.dll',
    'zlib1.dll',
    'libzlib.dll',
    'zlib.dll'
]

# 收集所有存在的 DLL 文件
binaries = []
for dll in core_dlls:
    dll_path = os.path.join(poppler_path, dll)
    if os.path.exists(dll_path):
        binaries.append((dll_path, '.'))  # 放在根目录下
        log(f"找到依赖文件: {dll}")
    else:
        log(f"警告: 找不到依赖文件 {dll}")

# 添加整个 poppler 目录
binaries.append((poppler_path, 'poppler'))
log(f"添加 poppler 目录: {poppler_path}")

a = Analysis(
    ['pdf_merger.py'],
    pathex=[],
    binaries=binaries,
    datas=[('icon.png', '.')],
    hiddenimports=[
        'pkg_resources.py2_warn',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'notebook', 'scipy', 'pandas', 'numpy.random',
        'docutils', 'pycparser', 'psutil', 'PIL.ImageQt', 'PyQt5',
        'PyQt6', 'PySide2', 'PySide6', 'IPython', 'jedi', 'testpath',
        'nbconvert', 'nbformat', 'traitlets', 'pygments', 'sqlite3',
        'colorama', 'tornado', 'zmq', 'numpy.core._dotblas', 'PyInstaller',
        'win32com', 'win32api', 'win32gui', 'win32security', 'win32event',
        'win32process', 'win32pipe', 'win32file', 'pywintypes', 'pythoncom'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 移除不需要的二进件
a.binaries = [x for x in a.binaries if not x[0].startswith(('mfc', 'vcruntime', 'api-ms-win', 'qt', 'PyQt', 'sip'))]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PDF合并工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[
        'vcruntime140.dll',
        'python*.dll',
        'api-ms-win*.dll',
        'ucrtbase.dll',
        'VCRUNTIME140.dll',
        'msvcp140.dll',
        'python3*.dll',
        '_ssl.pyd',
        '_socket.pyd',
        'unicodedata.pyd',
        'libcrypto*.dll',
        'libssl*.dll',
        'tcl86t.dll',
        'tk86t.dll',
    ],
    upx_args=['--best', '--lzma', '--ultra-brute'],
    runtime_tmpdir=None,
    console=False,  # 改为 False 以隐藏控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ico_path,
    uac_admin=False,
    uac_uiaccess=False,
    win_private_assemblies=True,
    win_no_prefer_redirects=True,
)

log("打包完成，临时文件保留在: " + ico_path)
log("程序执行完毕，按任意键退出...")