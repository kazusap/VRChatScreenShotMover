using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

using System.Windows.Media.Imaging;
using System.IO;
using System.Drawing;

namespace VRCSCMV
{
    class Program
    {
        static void Main(string[] args)
        {
            //VRChat_1920x1080_2019-09-27_01-30-32.838.png

            //VRChat_2022-09-10_00-27-20.675_1920x1080.png
            // 引数の数チェック
            Debug.WriteLine(args.Length.ToString());
            if (args.Length != 3)
            {
                #region 引数が規定の数以外のときの表示
                //バージョンの取得
                System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
                System.Version ver = asm.GetName().Version;

                Console.WriteLine("\r\nVRCSCMV(VRChat ScreenShot MoVer) ver " + ver.ToString() + " by 高槻かずさ" );
                Console.Write("\r\n");
                Console.WriteLine("■説明\r\nVRChatのスクリーンショットをjpgに変換してexifに撮影日を付加、移動します。");
                Console.WriteLine("※注意：元のPNGファイルは削除されます。");
                Console.Write("\r\n");
                Console.WriteLine("Usage : VRCSCMV.exe [VRChatスクリーンショット保存フォルダ] [移動先フォルダ] [n時までを前日扱いとする]");
                Console.WriteLine(" 例 : VRCSCMV C:\\Users\\username\\Pictures\\VRChat C:\\MovedPics 6");
                Console.WriteLine("     [C:\\Users\\username\\Pictures\\VRChat] から [C:\\MovedPics] へ移動。 06:00までを前日扱いとする。");
#if DEBUG
            Console.ReadKey();
#endif
                return;
                #endregion
            }

            // 引数0 = プログラムファイル名
            // 引数1 = VRChatスクリーンショットフォルダ
            // 引数2 = 移動先フォルダ（この下に日付フォルダを作成して移動）
            // 引数3 = 前日扱いとする（6にすると06:00までは前日と認識）
            Debug.WriteLine("args[0]:" + args[0]);
            Debug.WriteLine("args[1]:" + args[1]);
            Debug.WriteLine("args[2]:" + args[2]);
            //Debug.WriteLine("args[3]:" + args[3]);

            // VRChatスクリーンショットフォルダ
            string strFolder = args[0];

            // 移動先フォルダ（この下に日付フォルダを作成して移動）
            string strDestFolder = args[1];

            // 前日扱い時間
            int timeSpan_hour;
            if (int.TryParse(args[2], out timeSpan_hour))
            {
                timeSpan_hour = timeSpan_hour * -1;
            }
            else
            {
                Console.WriteLine("ERROR:時間は[0～23]にしてね。");
                return;
            }

            // サブディレクトリ一覧を取得
//            IEnumerable<string> dirs = System.IO.Directory.EnumerateDirectories(strFolder);


            // スクリーンショットのファイル一覧を取得(PNG)
            IEnumerable<string> files = System.IO.Directory.EnumerateFiles(strFolder, "*.PNG", System.IO.SearchOption.AllDirectories);

            // ファイルカウント
            int iSuccessCount = 0;

            try
            {
                // PNGファイルごとの処理
                foreach (string fullpath in files)
                {
                    // ファイル名のみ取得
                    string fname = System.IO.Path.GetFileNameWithoutExtension(fullpath);

                    Debug.WriteLine(fname);
                    Console.Write("filename " + fname + " ... ");

                    // ファイル名分割
                    string[] subname = fname.Split('_');

                    // 0        1            2              3
                    // [VRChat]_[2022-09-10]_[00-27-20.675]_[1920x1080.png]

                    // 日時文字列取得    YYYY-MM-DD                           HH-MI-HH.FFF
                    string strDatetime = subname[1].Replace('-', '/') + ' ' + subname[2].Replace('-', ':');

                    // 日時クラスに変換
                    DateTime dt = DateTime.Parse(strDatetime);

                    // 指定時間ずらす
                    DateTime dtOffset = dt.Add(TimeSpan.FromHours(timeSpan_hour));

                    //日付文字列取得
                    string strFoderDate = dtOffset.ToString("yyyy-MM-dd");

                    //日付フォルダが無ければ作成
                    string strDestPath = strDestFolder + @"\" + strFoderDate;
                    System.IO.Directory.CreateDirectory(strDestPath);

                    Console.Write("makeJPEG/");
                    //JPGを新規作成
                    string jpgpath = ConvertToJpeg(fullpath, 100);

                    bool flg = false;
                    while (!flg)
                    {
                        try
                        {
                            Console.Write("AddExif/");
                            //Exifに撮影日を付加
                            AddExifData(jpgpath, jpgpath, dt);
                            flg = true;
                        }
                        catch (IOException ex)
                        {
                            System.Threading.Thread.Sleep(100);
                        }
                    }

                    Console.Write("Move/");
                    //ファイル(JPG)を移動
                    string strDestFilePath = strDestPath + @"\" + System.IO.Path.GetFileName(jpgpath);
                    System.IO.File.Move(jpgpath, strDestFilePath);

                    Console.Write("DeleteOrigin/ ");
                    //元ファイル(PNG)を削除
                    System.IO.File.Delete(fullpath);

                    Console.Write("done.\r\n");

                    iSuccessCount++;

                }
            }
            catch (Exception ignore)
            {
                Console.WriteLine("\r\nエラー：" + ignore.Message);
                Console.WriteLine("中断しました。");
            }

            Console.WriteLine("Moved " + iSuccessCount.ToString() + " files.");

#if DEBUG
            Console.ReadKey();
#endif
        }

        /// <summary>
        /// JPGファイルに変換
        /// </summary>
        /// <param name="pngPath">元のPNGファイルフルパス</param>
        /// <param name="quality">品質(0-100)</param>
        /// <returns></returns>
        static string ConvertToJpeg(string pngPath, int quality)
        {
            string jpgPath = null;

            using (Bitmap bmp = new Bitmap(pngPath))
            {
                try
                {
                    string ext = System.IO.Path.GetExtension(pngPath);
                    jpgPath = pngPath.Replace(ext, "") + ".jpg";

                    System.Drawing.Imaging.EncoderParameters eps = new System.Drawing.Imaging.EncoderParameters(1);
                    System.Drawing.Imaging.EncoderParameter ep =new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)quality);
                    eps.Param[0] = ep;

                    System.Drawing.Imaging.ImageCodecInfo ici = GetEncoderInfo("image/jpeg");

                    bmp.Save(jpgPath, ici, eps);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                bmp.Dispose();
            }

            return jpgPath;
        }

        /// <summary>
        /// MimeTypeで指定されたImageCodecInfoを探して返す
        /// </summary>
        /// <param name="mineType">MimeType文字列</param>
        /// <returns></returns>
        private static System.Drawing.Imaging.ImageCodecInfo GetEncoderInfo(string mineType)
        {
            //GDI+ に組み込まれたイメージ エンコーダに関する情報をすべて取得
            System.Drawing.Imaging.ImageCodecInfo[] encs = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();

            //指定されたMimeTypeを探して見つかれば返す
            foreach (System.Drawing.Imaging.ImageCodecInfo enc in encs)
            {
                if (enc.MimeType == mineType)
                {
                    return enc;
                }
            }
            return null;
        }

        /// <summary>
        /// Exifデータを付加（撮影日）
        /// </summary>
        /// <param name="sourcePath">ソースjpgファイルパス</param>
        /// <param name="outcomePath">書き出し先ファイルパス</param>
        /// <param name="date">撮影日</param>
        static void AddExifData(string sourcePath, string outcomePath, DateTime date)
        {
            byte[] buff;

            using (var sourceStream = File.Open(sourcePath, FileMode.Open))
            using (var ms = new MemoryStream())
            {
                sourceStream.CopyTo(ms);

                buff = EditDateTaken(ms, date);
            }

            if (buff == null)
                return;

            using (var outcomeStream = File.Create(outcomePath))
            {
                outcomeStream.Write(buff, 0, buff.Length);
            }

        }

        /// <summary>
        /// Query paths for padding
        /// </summary>
        private static readonly List<string> queryPadding = new List<string>()
{
    "/app1/ifd/PaddingSchema:Padding", // Query path for IFD metadata
    "/app1/ifd/exif/PaddingSchema:Padding", // Query path for EXIF metadata
    "/xmp/PaddingSchema:Padding", // Query path for XMP metadata
};

        /// <summary>
        /// Edit date taken field of Exif metadata of image data.
        /// </summary>
        /// <param name="source">Stream of source image data in JPG format</param>
        /// <param name="date">Date to be set</param>
        /// <returns>Byte array of outcome image data</returns>
        public static Byte[] EditDateTaken(Stream source, DateTime date)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            if (date == null)
                throw new ArgumentNullException("date");

            if (0 < source.Position)
                source.Seek(0, SeekOrigin.Begin);

            // Create BitmapDecoder for a lossless transcode.
            var sourceDecoder = BitmapDecoder.Create(
                source,
                BitmapCreateOptions.PreservePixelFormat | BitmapCreateOptions.IgnoreColorProfile,
                BitmapCacheOption.None);

            // Check if the source image data is in JPG format.
            if (!sourceDecoder.CodecInfo.FileExtensions.Contains("jpg"))
                return null;

            if ((sourceDecoder.Frames[0] == null) || (sourceDecoder.Frames[0].Metadata == null))
                return null;

            var sourceMetadata = sourceDecoder.Frames[0].Metadata.Clone() as BitmapMetadata;

            // Add padding (4KiB) to metadata.
            queryPadding.ForEach(x => sourceMetadata.SetQuery(x, 4096U));

            using (var ms = new MemoryStream())
            {
                // Perform a lossless transcode with metadata which includes added padding.
                var outcomeEncoder = new JpegBitmapEncoder();

                outcomeEncoder.Frames.Add(BitmapFrame.Create(
                    sourceDecoder.Frames[0],
                    sourceDecoder.Frames[0].Thumbnail,
                    sourceMetadata,
                    sourceDecoder.Frames[0].ColorContexts));

                outcomeEncoder.Save(ms);

                // Create InPlaceBitmapMetadataWriter.
                ms.Seek(0, SeekOrigin.Begin);

                var outcomeDecoder = BitmapDecoder.Create(ms, BitmapCreateOptions.None, BitmapCacheOption.Default);

                var metadataWriter = outcomeDecoder.Frames[0].CreateInPlaceBitmapMetadataWriter();

                // Edit date taken field by accessing property of InPlaceBitmapMetadataWriter.
                metadataWriter.DateTaken = date.ToString();

                // Edit date taken field by using query with path string.
                metadataWriter.SetQuery("/app1/ifd/exif/{ushort=36867}", date.ToString("yyyy:MM:dd HH:mm:ss"));

                // Try to save edited metadata to stream.
                if (metadataWriter.TrySave())
                {
                    Debug.WriteLine("InPlaceMetadataWriter succeeded!");
                    return ms.ToArray();
                }
                else
                {
                    Debug.WriteLine("InPlaceMetadataWriter failed!");
                    return null;
                }
            }
        }



    }
}
