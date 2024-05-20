# Blazor FileViewer 文件预览 组件  

### 目前支支持 Excel(.docx) 和 Word(.xlsx) 格式

示例:

https://www.blazor.zone/fileViewers

https://blazor.app1.es/fileViewers

使用方法:

1.nuget包

```BootstrapBlazor.FileViewer```

2._Imports.razor 文件 或者页面添加 添加组件库引用

```@using BootstrapBlazor.Components```


3.razor页面
```
<FileViewer Filename="c:/DemoShared/sample.xlsx" />

<FileViewer Filename="c:/DemoShared/sample.docx" />

<FileViewer Filename="https://localhost:5011/_content/DemoShared/sample.xlsx" />

<FileViewer Filename="https://localhost:5011/_content/DemoShared/sample.docx" />

<FileViewer @ref="fileViewer" Filename=@Url />

@code{
    private string Url { get; set; } = ("c:/sample.docx");

    private async Task Apply()
    {
        await fileViewer.Reload(Url);
    }
}
```

4.参数说明

|  参数   | 说明  | 默认值  | 
|  ----  | ----  | ----  | 
| Filename  | Excel/Word 文件路径或者URL |  | 
| Width  | 宽度 | 100% | 
| Height  | 高度 | 700px | 
| StyleString  | 组件外观 Css Style |  | 
| Html  | 设置 Html 直接渲染  |  | 
| Stream  | 用于渲染的文件流,为空则用Filename参数读取文件 | null | 
| IsExcel  | 文件流模式需要指定是否 Excel  | false | 
| NodataString  | 无数据提示文本 | 无数据 | 
| LoadingString  | 载入中提示文本  | 载入中... | 
| Reload(string filename) | 重新载入文件方法 | |
| Reload(Stream stream) | 重新载入流方法 | |
| Refresh() | 刷新方法 | |

 
### 5.特别说明

如果在 Linux 下使用需要安装 libgdiplus 并开启 System.Drawing support.

相关错误提示: The type initializer for 'Gdip' threw an exception.

Enable System.Drawing support for non-Windows platforms: (reference):

In your project file (*.csproj), add:
```
 <ItemGroup>
     <RuntimeHostConfigurationOption Include="System.Drawing.EnableUnixSupport" Value="true" />
 </ItemGroup>
```

---
#### 更新历史

v8.0.2
- 修复 [依赖包日期格式转换错误](https://github.com/densen2014/BootstrapBlazor.FileViewer/issues/4)
- 修复 [组件依赖](https://github.com/densen2014/BootstrapBlazor.FileViewer/issues/2)

v7.0.2
- 修复 [预览表格时间转换错误](https://github.com/densen2014/BootstrapBlazor.FileViewer/issues/1)

v7.0.3
- 添加 Reload(Stream stream) : 重新载入流方法
- 修复 Reload(string filename) 不清空 Stream

---
#### Blazor 组件

[条码扫描 ZXingBlazor](https://www.nuget.org/packages/ZXingBlazor#readme-body-tab)
[![nuget](https://img.shields.io/nuget/v/ZXingBlazor.svg?style=flat-square)](https://www.nuget.org/packages/ZXingBlazor) 
[![stats](https://img.shields.io/nuget/dt/ZXingBlazor.svg?style=flat-square)](https://www.nuget.org/stats/packages/ZXingBlazor?groupby=Version)

[图片浏览器 Viewer](https://www.nuget.org/packages/BootstrapBlazor.Viewer#readme-body-tab)

[手写签名 SignaturePad](https://www.nuget.org/packages/BootstrapBlazor.SignaturePad#readme-body-tab)

[定位/持续定位 Geolocation](https://www.nuget.org/packages/BootstrapBlazor.Geolocation#readme-body-tab)

[屏幕键盘 OnScreenKeyboard](https://www.nuget.org/packages/BootstrapBlazor.OnScreenKeyboard#readme-body-tab)

[百度地图 BaiduMap](https://www.nuget.org/packages/BootstrapBlazor.BaiduMap#readme-body-tab)

[谷歌地图 GoogleMap](https://www.nuget.org/packages/BootstrapBlazor.Maps#readme-body-tab)

[蓝牙和打印 Bluetooth](https://www.nuget.org/packages/BootstrapBlazor.Bluetooth#readme-body-tab)

[PDF阅读器 PdfReader](https://www.nuget.org/packages/BootstrapBlazor.PdfReader#readme-body-tab)

[文件系统访问 FileSystem](https://www.nuget.org/packages/BootstrapBlazor.FileSystem#readme-body-tab)

[光学字符识别 OCR](https://www.nuget.org/packages/BootstrapBlazor.OCR#readme-body-tab)

[电池信息/网络信息 WebAPI](https://www.nuget.org/packages/BootstrapBlazor.WebAPI#readme-body-tab)

[视频播放器 VideoPlayer](https://www.nuget.org/packages/BootstrapBlazor.VideoPlayer#readme-body-tab)

[文件预览 FileViewer](https://www.nuget.org/packages/BootstrapBlazor.FileViewer#readme-body-tab)

[视频播放器 VideoPlayer](https://www.nuget.org/packages/BootstrapBlazor.VideoPlayer#readme-body-tab)

[图像裁剪 ImageCropper](https://www.nuget.org/packages/BootstrapBlazor.ImageCropper#readme-body-tab)

[视频播放器 BarcodeGenerator](https://www.nuget.org/packages/BootstrapBlazor.BarcodeGenerator#readme-body-tab)

#### AlexChow

[今日头条](https://www.toutiao.com/c/user/token/MS4wLjABAAAAGMBzlmgJx0rytwH08AEEY8F0wIVXB2soJXXdUP3ohAE/?) | [博客园](https://www.cnblogs.com/densen2014) | [知乎](https://www.zhihu.com/people/alex-chow-54) | [Gitee](https://gitee.com/densen2014) | [GitHub](https://github.com/densen2014)

