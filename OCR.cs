using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Windows.Media.Imaging;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using System.Threading.Tasks;
using System.Drawing;

namespace Produire.PImaging
{
	public class OCR : IProduireStaticClass
	{
		/// <summary>
		/// 画像から文字認識します。
		/// </summary>
		/// <param name="image">読み取り画像</param>
		/// <param name="lang">対象の言語</param>
		/// <returns></returns>
		[自分で]
		public string 読み取る([を]Image image, [として, 省略]string lang)
		{
			Task<SoftwareBitmap> t1 = null;
			Task.Run(() =>
			{
				t1 = ConvertSoftwareBitmap((Bitmap)image);
			}).Wait();
			SoftwareBitmap x1 = t1.Result;

			if (string.IsNullOrEmpty(lang)) lang = "ja-JP";

			Task<OcrResult> t2 = null;
			Task.Run(() =>
			{
				t2 = RunOcr(x1, lang);
			}).Wait();
			OcrResult x2 = t2.Result;
			StringBuilder builder = new StringBuilder();
			foreach (var line in x2.Lines)
			{
				if (builder.Length > 0) builder.AppendLine();
				var words = line.Words;
				bool lastHalf = false;
				for (int i = 0; i < words.Count; i++)
				{
					var word = words[i];
					bool isHalf = IsHalf(word.Text);
					if (lastHalf != isHalf)
					{
						if (i > 0) builder.Append(" ");
					}
					builder.Append(word.Text);
				}
			}
			return builder.ToString();
		}

		private bool IsHalf(string text)
		{
			bool half = true;
			for (int i = 0; i < text.Length; i++)
			{
				char c = text[i];
				if (char.IsDigit(c) || char.IsLower(c) || char.IsUpper(c) || char.IsWhiteSpace(c))
				{
				}
				else
					return false;
			}
			return half;
		}

		private async Task<SoftwareBitmap> ConvertSoftwareBitmap(Bitmap bitmap)
		{
			System.Windows.Media.Imaging.BitmapFrame bitmapSource = null;
			using (var wfStream = new MemoryStream())
			{
				bitmap.Save(wfStream, System.Drawing.Imaging.ImageFormat.Bmp);
				wfStream.Seek(0, SeekOrigin.Begin);
				bitmapSource = System.Windows.Media.Imaging.BitmapFrame.Create(wfStream, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
			}

			SoftwareBitmap sbitmap = null;

			using (MemoryStream wpfStream = new MemoryStream())
			{
				var encoder = new BmpBitmapEncoder();
				encoder.Frames.Add(bitmapSource);
				encoder.Save(wpfStream);

				var irstream = WindowsRuntimeStreamExtensions.AsRandomAccessStream(wpfStream);

				var decorder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(irstream);
				sbitmap = await decorder.GetSoftwareBitmapAsync();
			}

			return sbitmap;
		}

		private async Task<OcrResult> RunOcr(SoftwareBitmap sbitmap, string lang)
		{
			OcrEngine engine = OcrEngine.TryCreateFromLanguage(new Windows.Globalization.Language(lang));
			var result = await engine.RecognizeAsync(sbitmap);
			return result;
		}

	}
}
