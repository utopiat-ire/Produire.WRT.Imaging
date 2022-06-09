using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using Windows.Devices.Enumeration;
using Windows.Devices.Scanners;
using Windows.Graphics.Imaging;
using Windows.Storage;

namespace Produire.PImaging
{
	public class スキャナ : IProduireStaticClass
	{
		ImageScanner myScanner;
		CancellationTokenSource cancellationToken;

		DeviceWatcher scannerWatcher;
		List<DeviceInformation> scannerList = new List<DeviceInformation>();

		#region 手順

		/// <summary>
		/// 利用できるイメージスキャナを列挙します
		/// </summary>
		[自分を]
		public void 列挙する()
		{
			scannerWatcher = DeviceInformation.CreateWatcher(DeviceClass.ImageScanner);
			scannerWatcher.Added += OnScannerAdded;
			scannerWatcher.Removed += OnScannerRemoved;
			scannerWatcher.EnumerationCompleted += OnScannerEnumerationComplete;
			scannerWatcher.Start();
			for (; ; )
			{
				if (scannerWatcher.Status == DeviceWatcherStatus.EnumerationCompleted
						|| scannerWatcher.Status == DeviceWatcherStatus.Stopped) break;
				Task t1 = null;
				Task.Run(() =>
				{
					t1 = Task.Delay(500);
				}).Wait();
			}
		}

		private void OnScannerAdded(DeviceWatcher sender, DeviceInformation deviceInfo)
		{
			foreach (var item in scannerList)
			{
				if (item.Id == deviceInfo.Id)
				{
					return;
				}
			}
			if (deviceInfo != null) scannerList.Add(deviceInfo);
		}

		private void OnScannerRemoved(DeviceWatcher sender, DeviceInformationUpdate args)
		{
			int i = 0;
			for (; i < scannerList.Count; i++)
			{
				if (scannerList[i].Id == args.Id)
				{
					scannerList.RemoveAt(i);
					break;
				}
			}
		}
		private void OnScannerEnumerationComplete(DeviceWatcher sender, object args)
		{
			scannerWatcher.Stop();
		}

		/// <summary>
		/// 読み取りに使用するイメージスキャナを選択します。
		/// </summary>
		/// <param name="デバイス名"></param>
		/// <returns></returns>
		[自分へ]
		public bool 選択する([を]string デバイス名)
		{
			if (string.IsNullOrEmpty(デバイス名)) throw new ProduireException("スキャナのデバイスIDまたは名称を指定してください。");
			string deviceId = null;
			foreach (var item in scannerList)
			{
				if (item.Id == デバイス名)
				{
					deviceId = item.Id;
					break;
				}
			}
			if (deviceId == null)
			{
				foreach (var item in scannerList)
				{
					if (item.Name == デバイス名)
					{
						deviceId = item.Id;
						break;
					}
				}
			}
			if (deviceId == null)
			{
				throw new ProduireException("該当するスキャナが見つかりません。" + デバイス名);
			}

			Task<ImageScanner> t1 = null;
			Task.Run(() =>
			{
				try
				{
					t1 = ImageScanner.FromIdAsync(deviceId).AsTask<ImageScanner>();
				}
				catch (Exception ex)
				{
					if (ex.InnerException != null) throw ex.InnerException;
					throw;
				}
			}).Wait();
			myScanner = t1.Result;

			return myScanner != null;
		}

		/// <summary>
		/// イメージスキャナから画像を読み取って指定したフォルダへ保存します。
		/// </summary>
		/// <param name="パス">保存先のフォルダ</param>
		/// <returns>保存した画像ファイル</returns>
		[自分から]
		public string[] 読み取る([へ, 省略]string パス)
		{
			if (myScanner == null)
			{
				string deviceId = GetDefaultDevice();
				選択する(deviceId);
			}

			cancellationToken = new CancellationTokenSource();

			ImageScannerScanResult result;
			Task<ImageScannerScanResult> t1 = null;
			Task.Run(() =>
			{
				t1 = ScanFilesToFolder();
			}).Wait();
			result = t1.Result;

			if (!string.IsNullOrEmpty(パス) && !Directory.Exists(パス)) Directory.CreateDirectory(パス);
			List<string> list = new List<string>();
			foreach (StorageFile file in result.ScannedFiles)
			{
				string path = Path.Combine(パス, Path.GetFileName(file.Path));
				if (file.Path != path) File.Move(file.Path, path);
				list.Add(path);
			}

			return list.ToArray();
		}

		private async Task<ImageScannerScanResult> ScanFilesToFolder()
		{
			StorageFolder folder = KnownFolders.PicturesLibrary;
			if (folder == null) return null;

			ImageScannerScanSource source = ImageScannerScanSource.AutoConfigured;
			if (selectedSource != -1)
			{
				source = 読み取り元;
				if (!myScanner.IsScanSourceSupported(source))
				{
					throw new ProduireException("この読み取り元はサポートされていません。");
				}
			}
			else if (myScanner.IsScanSourceSupported(ImageScannerScanSource.AutoConfigured))
			{
				source = ImageScannerScanSource.AutoConfigured;
			}
			else
			{
				source = ImageScannerScanSource.Default;
			}

			return await myScanner.ScanFilesToFolderAsync(source, folder).AsTask(cancellationToken.Token);
		}

		/// <summary>
		/// イメージスキャナからプレビュー用の画像を読み取り画像として返します。
		/// </summary>
		/// <returns>画像</returns>
		[自分から]
		public Image プレビュー()
		{
			if (myScanner == null)
			{
				string deviceId = GetDefaultDevice();
				選択する(deviceId);
			}

			cancellationToken = new CancellationTokenSource();

			Image result;

			Task<Image> t1 = null;
			Task.Run(() =>
			{
				t1 = SavePreview();
			}).Wait();
			try
			{
				result = t1.Result;
			}
			catch (Exception ex)
			{
				if (ex.InnerException != null) throw ex.InnerException;
				throw;
			}

			return result;
		}

		private async Task<Image> SavePreview()
		{
			Image image = null;
			ImageScannerPreviewResult result = null;
			using (MemoryStream netStream = new MemoryStream())
			{
				var scanStream = WindowsRuntimeStreamExtensions.AsRandomAccessStream(netStream);

				ImageScannerScanSource source = ImageScannerScanSource.AutoConfigured;
				if (selectedSource != -1)
				{
					source = 読み取り元;
					if (!myScanner.IsScanSourceSupported(source))
					{
						throw new ProduireException("この読み取り元はサポートされていません。");
					}
				}
				else if (myScanner.IsScanSourceSupported(ImageScannerScanSource.AutoConfigured))
				{
					source = ImageScannerScanSource.AutoConfigured;
				}
				else
				{
					source = ImageScannerScanSource.Default;
				}
				result = await myScanner.ScanPreviewToStreamAsync(source, scanStream);
				await scanStream.FlushAsync();

				scanStream.Seek(0);

				var decoder = await BitmapDecoder.CreateAsync(scanStream);
				var data = await decoder.GetPixelDataAsync();

				using (MemoryStream netStream2 = new MemoryStream())
				{
					var bmpStream = WindowsRuntimeStreamExtensions.AsRandomAccessStream(netStream2);
					var encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, bmpStream);

					encoder.SetPixelData(
						BitmapPixelFormat.Bgra8, BitmapAlphaMode.Straight,
						(uint)decoder.PixelWidth, (uint)decoder.PixelHeight,
						decoder.DpiX, decoder.DpiY, data.DetachPixelData());

					encoder.BitmapTransform.Bounds = new BitmapBounds() { Width = decoder.PixelWidth, Height = decoder.PixelHeight };

					await encoder.FlushAsync();

					bmpStream.Seek(0);
					image = Image.FromStream(netStream2);
				}
			}
			return image;
		}

		private string GetDefaultDevice()
		{
			if (scannerList.Count == 0) 列挙する();
			foreach (var item in scannerList)
			{
				if (item.IsDefault) return item.Id;
			}
			if (scannerList.Count > 0) return scannerList[0].Id;
			return null;
		}

		public void CancelScanning()
		{
			if (cancellationToken != null)
			{
				cancellationToken.Cancel();
			}
		}

		#endregion

		#region 設定項目

		/// <summary>
		/// 利用できるイメージスキャナのデバイス名一覧を返します
		/// </summary>
		public string[] 一覧
		{
			get
			{
				if (scannerList.Count == 0) 列挙する();
				List<string> list = new List<string>();
				foreach (var item in scannerList)
				{
					list.Add(item.Name);
				}
				return list.ToArray();
			}
		}
		/// <summary>
		/// 利用できるイメージスキャナのデバイスID一覧を返します
		/// </summary>
		public string[] デバイスID一覧
		{
			get
			{
				if (scannerList.Count == 0) 列挙する();
				List<string> list = new List<string>();
				foreach (var item in scannerList)
				{
					list.Add(item.Id);
				}
				return list.ToArray();
			}
		}

		int selectedSource = -1;
		/// <summary>
		/// 読み取りに使用するソース
		/// </summary>
		public ImageScannerScanSource 読み取り元
		{
			get { return selectedSource == -1 ? ImageScannerScanSource.AutoConfigured : (ImageScannerScanSource)selectedSource; }
			set { selectedSource = (int)value; }
		}

		#endregion
	}

	[列挙体(typeof(ImageScannerScanSource))]
	public enum 読み取り元列挙
	{
		既定 = ImageScannerScanSource.Default,
		フィーダ = ImageScannerScanSource.Feeder,
		フラットヘッド = ImageScannerScanSource.Flatbed,
		自動 = ImageScannerScanSource.AutoConfigured
	}
}
