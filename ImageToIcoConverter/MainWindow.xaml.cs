using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media.Imaging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;

namespace ImageToIcoConverter
{
    public sealed partial class MainWindow : Window
    {
        private StorageFile? _selectedImageFile;
        private List<Size> _selectedIcoSizes = [];
        private MemoryStream? _icoMemoryStream;

        public MainWindow()
        {
            this.InitializeComponent();
            ExtendsContentIntoTitleBar = true;
            Cb16x16.IsChecked = true;
            Cb32x32.IsChecked = true;
            Cb48x48.IsChecked = true;
            UpdateSelectedSizes();
        }

        #region 按钮事件
        private async void BtnSelectImage_Click(object sender, RoutedEventArgs e)
        {
            var filePicker = new FileOpenPicker();
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(filePicker, hWnd);

            filePicker.FileTypeFilter.Add(".png");
            filePicker.FileTypeFilter.Add(".jpg");
            filePicker.FileTypeFilter.Add(".jpeg");
            filePicker.FileTypeFilter.Add(".bmp");
            filePicker.FileTypeFilter.Add(".gif");
            filePicker.FileTypeFilter.Add(".tiff");

            var file = await filePicker.PickSingleFileAsync();
            if (file != null)
            {
                _selectedImageFile = file;
                await LoadOriginalImageAsync(file);
                await GenerateIcoAndPreviewAsync();
                BtnSaveIco.IsEnabled = true;
            }
        }

        private async void BtnSaveIco_Click(object sender, RoutedEventArgs e)
        {
            if (_icoMemoryStream == null || _icoMemoryStream.Length == 0)
            {
                ShowToast("请先选择图片并生成ICO");
                return;
            }

            var savePicker = new FileSavePicker();
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(savePicker, hWnd);

            savePicker.SuggestedFileName = "converted_icon";
            savePicker.FileTypeChoices.Add("ICO文件", new List<string> { ".ico" });
            savePicker.DefaultFileExtension = ".ico";

            var saveFile = await savePicker.PickSaveFileAsync();
            if (saveFile != null)
            {
                using (var stream = await saveFile.OpenStreamForWriteAsync())
                {
                    _icoMemoryStream.Seek(0, SeekOrigin.Begin);
                    await _icoMemoryStream.CopyToAsync(stream);
                    await stream.FlushAsync();
                }
                ShowToast("ICO保存成功！");
            }
        }
        #endregion

        #region 尺寸选择
        private void IcoSize_Checked(object sender, RoutedEventArgs e)
        {
            UpdateSelectedSizes();
            if (_selectedImageFile != null) _ = GenerateIcoAndPreviewAsync();
        }

        private void IcoSize_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateSelectedSizes();
            if (_selectedImageFile != null) _ = GenerateIcoAndPreviewAsync();
        }

        private void UpdateSelectedSizes()
        {
            _selectedIcoSizes.Clear();
            if (Cb16x16.IsChecked == true) _selectedIcoSizes.Add(new Size(16, 16));
            if (Cb32x32.IsChecked == true) _selectedIcoSizes.Add(new Size(32, 32));
            if (Cb48x48.IsChecked == true) _selectedIcoSizes.Add(new Size(48, 48));
            if (Cb64x64.IsChecked == true) _selectedIcoSizes.Add(new Size(64, 64));
            if (Cb128x128.IsChecked == true) _selectedIcoSizes.Add(new Size(128, 128));
            if (Cb256x256.IsChecked == true) _selectedIcoSizes.Add(new Size(256, 256));
        }
        #endregion

        #region 图片加载
        private async System.Threading.Tasks.Task LoadOriginalImageAsync(StorageFile file)
        {
            try
            {
                using (var stream = await file.OpenAsync(FileAccessMode.Read))
                {
                    var bitmap = new BitmapImage();
                    await bitmap.SetSourceAsync(stream);
                    ImgOriginal.Source = bitmap;
                    TxtNoOriginal.Visibility = Visibility.Collapsed;
                }
            }
            catch (Exception ex)
            {
                ShowToast($"加载图片失败：{ex.Message}");
            }
        }
        #endregion

        #region 修复后的ICO生成核心逻辑
        private async System.Threading.Tasks.Task GenerateIcoAndPreviewAsync()
        {
            if (_selectedImageFile == null || _selectedIcoSizes.Count == 0)
            {
                ClearIcoPreview();
                return;
            }

            try
            {
                using (var fileStream = await _selectedImageFile.OpenStreamForReadAsync())
                {
                    using (var originalImage = System.Drawing.Image.FromStream(fileStream))
                    {
                        _icoMemoryStream = new MemoryStream();
                        // 使用修复后的ICO生成方法
                        CreateValidIcoFile(originalImage, _selectedIcoSizes, _icoMemoryStream);
                        await ShowIcoPreviewAsync(_icoMemoryStream);
                    }
                }
            }
            catch (Exception ex)
            {
                ShowToast($"生成ICO失败：{ex.Message}");
                ClearIcoPreview();
            }
        }

        // 严格遵循ICO格式规范生成ICO文件
        private void CreateValidIcoFile(System.Drawing.Image originalImage, List<Size> sizes, MemoryStream outputStream)
        {
            if (sizes == null || sizes.Count == 0)
                throw new ArgumentException("至少选择一个ICO尺寸");

            using (var writer = new BinaryWriter(outputStream, System.Text.Encoding.ASCII, true))
            {
                // --------------------------
                // 1. 写入ICO文件头 (6字节)
                // --------------------------
                writer.Write((short)0);          // Reserved (必须为0)
                writer.Write((short)1);          // Type: 1=ICO, 2=CUR
                writer.Write((short)sizes.Count); // 图像数量

                // --------------------------
                // 2. 准备所有图像数据并计算目录项
                // --------------------------
                var imageDataList = new List<byte[]>();
                var directoryEntries = new List<byte[]>();
                long currentOffset = 6 + (16 * sizes.Count); // 头 + 所有目录项的长度

                foreach (var size in sizes)
                {
                    // 调整图片尺寸
                    using (var resizedImage = new Bitmap(originalImage, size.Width, size.Height))
                    {
                        // 转换为32位ARGB（ICO要求带透明度）
                        byte[] bmpData = GetIcoCompatibleBmpData(resizedImage);
                        imageDataList.Add(bmpData);

                        // --------------------------
                        // 3. 写入目录项 (每个16字节)
                        // --------------------------
                        var dirEntry = new BinaryWriter(new MemoryStream());
                        // 宽度 (256时写0)
                        dirEntry.Write((byte)(size.Width == 256 ? 0 : size.Width));
                        // 高度 (256时写0，ICO格式中高度是实际的2倍，需注意)
                        dirEntry.Write((byte)(size.Height == 256 ? 0 : size.Height));
                        dirEntry.Write((byte)0); // 颜色数 (0=256色以上)
                        dirEntry.Write((byte)0); // Reserved (必须为0)
                        dirEntry.Write((short)1); // 颜色平面数 (必须为1)
                        dirEntry.Write((short)32); // 位深度 (32位ARGB)
                        dirEntry.Write((int)bmpData.Length); // 图像数据大小
                        dirEntry.Write((int)currentOffset); // 图像数据偏移量

                        directoryEntries.Add(((MemoryStream)dirEntry.BaseStream).ToArray());
                        currentOffset += bmpData.Length;
                        dirEntry.Close();
                    }
                }

                // --------------------------
                // 4. 写入所有目录项
                // --------------------------
                foreach (var entry in directoryEntries)
                {
                    writer.Write(entry);
                }

                // --------------------------
                // 5. 写入所有图像数据
                // --------------------------
                foreach (var data in imageDataList)
                {
                    writer.Write(data);
                }

                writer.Flush();
            }
        }

        // 转换为ICO兼容的BMP数据（去掉BMP文件头，调整行序）
        private byte[] GetIcoCompatibleBmpData(Bitmap bitmap)
        {
            // 锁定位图数据
            BitmapData bmpData = bitmap.LockBits(
                new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                ImageLockMode.ReadOnly,
                PixelFormat.Format32bppArgb);

            // 计算像素数据长度
            int byteCount = bmpData.Stride * bitmap.Height;
            byte[] pixelData = new byte[byteCount];
            System.Runtime.InteropServices.Marshal.Copy(bmpData.Scan0, pixelData, 0, byteCount);
            bitmap.UnlockBits(bmpData);

            // ICO要求BMP行序反转（从下到上）
            byte[] reversedPixelData = ReverseBmpRows(pixelData, bitmap.Width, bitmap.Height, bmpData.Stride);

            // 创建BITMAPINFOHEADER (40字节)
            var bmpHeader = new BinaryWriter(new MemoryStream());
            bmpHeader.Write((int)40);               // 头大小
            bmpHeader.Write((int)bitmap.Width);      // 宽度
            bmpHeader.Write((int)(bitmap.Height * 2)); // 高度（ICO格式要求写2倍）
            bmpHeader.Write((short)1);              // 平面数
            bmpHeader.Write((short)32);             // 位深度
            bmpHeader.Write((int)0);                // 压缩方式
            bmpHeader.Write((int)reversedPixelData.Length); // 图像大小
            bmpHeader.Write((int)0);                // 水平分辨率
            bmpHeader.Write((int)0);                // 垂直分辨率
            bmpHeader.Write((int)0);                // 颜色数
            bmpHeader.Write((int)0);                // 重要颜色数

            // 合并头和像素数据
            byte[] headerBytes = ((MemoryStream)bmpHeader.BaseStream).ToArray();
            bmpHeader.Close();

            return headerBytes.Concat(reversedPixelData).ToArray();
        }

        // 反转BMP行序（ICO格式要求从下到上存储像素）
        private byte[] ReverseBmpRows(byte[] pixelData, int width, int height, int stride)
        {
            byte[] reversedData = new byte[pixelData.Length];
            int rowSize = stride;

            for (int y = 0; y < height; y++)
            {
                int sourceRow = (height - 1 - y) * rowSize;
                int destRow = y * rowSize;
                Array.Copy(pixelData, sourceRow, reversedData, destRow, rowSize);
            }

            return reversedData;
        }
        #endregion

        #region 预览和辅助方法
        private async System.Threading.Tasks.Task ShowIcoPreviewAsync(MemoryStream icoStream)
        {
            icoStream.Seek(0, SeekOrigin.Begin);

            using (var randomAccessStream = new InMemoryRandomAccessStream())
            {
                await randomAccessStream.WriteAsync(icoStream.ToArray().AsBuffer());
                randomAccessStream.Seek(0);

                var bitmap = new BitmapImage();
                await bitmap.SetSourceAsync(randomAccessStream);
                ImgIcoPreview.Source = bitmap;
                TxtNoIco.Visibility = Visibility.Collapsed;
            }
        }

        private void ClearIcoPreview()
        {
            ImgIcoPreview.Source = null;
            TxtNoIco.Visibility = Visibility.Visible;
            _icoMemoryStream = null;
            BtnSaveIco.IsEnabled = false;
        }

        private void ShowToast(string message)
        {
            var toast = new ContentDialog
            {
                Title = "提示",
                Content = message,
                CloseButtonText = "确定",
                XamlRoot = this.Content.XamlRoot
            };
            _ = toast.ShowAsync();
        }
        #endregion
    }
}