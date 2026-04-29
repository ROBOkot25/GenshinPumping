using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using ImageFormat = System.Drawing.Imaging.ImageFormat;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Tesseract;
using ClosedXML.Excel;

namespace ScreenshotOcrReader
{
    #region КЛАССЫ ДАННЫХ
    public class CharacterData
    {
        public string FileName { get; set; }
        public string CharacterName { get; set; }
        public string CharacterLevel { get; set; }
        public string WeaponName { get; set; }
        public string WeaponLevel { get; set; }
        public string HP { get; set; }
        public string Atk { get; set; }
        public string Def { get; set; }
        public string MasteryStihiy { get; set; }
        public string ArtifactSet { get; set; }

        // Новые поля для доп. характеристик
        public string CritRate { get; set; }
        public string CritDMG { get; set; }
        public string EnergyRecharge { get; set; }

        // Уровни талантов
        public string Talent1 { get; set; }  // ЛКМ
        public string Talent2 { get; set; }  // ПКМ / Е
        public string Talent3 { get; set; } // Ульта
    }
    #endregion

    #region PROGRAM
    class Program
    {
        #region ПЕРЕМЕННЫЕ
        private static readonly string ScreenshotsFolder = @"screenshots";
        private static readonly string TessDataPath = @"tessdata";
        private static readonly string DebugPath = @"debug_photo";
        private static readonly string ExcelFilePath = @"Геншин.xlsx";
        private static readonly string OcrLanguage = "rus";

        [DllImport("kernel32.dll")]
        private static extern bool SetStdHandle(int nStdHandle, IntPtr hHandle);
        private const int STD_ERROR_HANDLE = -12;

        // Основные регионы (скрин персонажа и оружия)
        private static readonly Rectangle NameRegion = new Rectangle(135, 25, 300, 52);
        private static readonly Rectangle LevelRegion = new Rectangle(1453, 205, 265, 43);
        private static readonly Rectangle HPRegion = new Rectangle(1500, 283, 330, 40);
        private static readonly Rectangle AtkRegion = new Rectangle(1500, 313, 330, 40);
        private static readonly Rectangle DefRegion = new Rectangle(1500, 350, 330, 40);
        private static readonly Rectangle MsRegion = new Rectangle(1500, 380, 330, 40);

        // CHECK-регионы (доп. характеристики) — для идентификации персонажа и извлечения значений
        private static readonly Rectangle HPRegionCHEAK = new Rectangle(245, 65, 1400, 55);
        private static readonly Rectangle AtkRegionCHEAK = new Rectangle(645, 130, 1000, 55);
        private static readonly Rectangle DefRegionCHEAK = new Rectangle(645, 195, 1000, 55);
        private static readonly Rectangle MsRegionCHEAK = new Rectangle(645, 260, 1000, 55);
        private static readonly Rectangle CritRateRegion = new Rectangle(645, 430, 1120, 55);
        private static readonly Rectangle CritDMGRegion = new Rectangle(645, 490, 720, 55);
        private static readonly Rectangle RechargeRegionCHEAK = new Rectangle(645, 675, 720, 55);

        private static readonly Rectangle NameWeaponRegion = new Rectangle(1445, 110, 371, 81);
        private static readonly Rectangle BaseAtkLabelRegion = new Rectangle(1465, 220, 250, 50);
        private static readonly Rectangle LevelWeaponRegion = new Rectangle(1450, 300, 170, 125);

        private static readonly Rectangle ArtifactBONUSRegion = new Rectangle(1440, 370, 360, 50);
        private static readonly Rectangle ArtifactSetRegion = new Rectangle(1450, 380, 360, 450);

        private static readonly Rectangle Talent1Region = new Rectangle(1620, 160, 90, 60);
        private static readonly Rectangle Talent2Region = new Rectangle(1620, 245, 90, 60);
        private static readonly Rectangle Talent3Region = new Rectangle(1620, 340, 90, 170);
        #endregion 

        #region ЛОГ
        private static void Log(string message, ConsoleColor color = ConsoleColor.Gray)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ResetColor();
        }
        #endregion

        #region ТОЧКА ВХОДА
        static void Main()
        {
            if (!Directory.Exists(ScreenshotsFolder))
            {
                Log("Папка не найдена: " + ScreenshotsFolder, ConsoleColor.Red);
                Console.ReadKey();
                return;
            }
            if (!Directory.Exists(TessDataPath))
            {
                Log("Папка не найдена: " + TessDataPath, ConsoleColor.Red);
                Console.ReadKey();
                return;
            }

            PrepareDebugFolder();

            var logFileStream = new FileStream("tesseract_log.txt", FileMode.Create, FileAccess.Write, FileShare.Read);
            var logWriter = new StreamWriter(logFileStream) { AutoFlush = true };
            Console.SetError(logWriter);
            SetStdHandle(STD_ERROR_HANDLE, logFileStream.SafeFileHandle.DangerousGetHandle());

            string[] imageFiles = Directory.GetFiles(ScreenshotsFolder, "*.*", SearchOption.TopDirectoryOnly);
            var supportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif"
            };

            // Два прохода:
            // 1й — все скриншоты кроме "Доп.характеристики" (там есть имя персонажа)
            // 2й — скриншоты "Доп.характеристики" (нужны уже собранные данные для матчинга)
            var mainScreenshots = new List<CharacterData>();
            var extraStatsScreenshots = new List<(string TempPath, string OrigPath, string FileName)>();

            using (var engine = new TesseractEngine(TessDataPath, OcrLanguage, EngineMode.Default))
            {
                engine.SetVariable("tessedit_pageseg_mode", 6);
                engine.SetVariable("tessedit_char_whitelist",
                    "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя0123456789/:%+ ");
                engine.SetVariable("classify_bln_numeric_mode", 1);
                engine.SetVariable("preserve_interword_spaces", "1");
                engine.SetVariable("textord_force_make_prop_words", "0");
                engine.SetVariable("tessedit_ocr_engine_mode", "1");
                engine.SetVariable("textord_min_linesize", "1.5");
                engine.SetVariable("edges_min_nonvertical", "0.5");

                // ── ПРОХОД 1: все обычные скриншоты ──────────────────────────────────
                foreach (string filePath in imageFiles)
                {
                    string ext = Path.GetExtension(filePath);
                    if (!supportedExtensions.Contains(ext)) continue;

                    string fileName = Path.GetFileName(filePath);
                    Log("\nОбрабатываю: " + fileName, ConsoleColor.Cyan);

                    try
                    {
                        using (var original = new Bitmap(filePath))
                        using (var enhanced = EnhanceImage(original))
                        {
                            // Сохраняем enhanced во временный файл
                            string tempPath = Path.GetTempFileName();
                            enhanced.Save(tempPath, ImageFormat.Png);

                            // Определяем тип скриншота — быстрый однопопыточный OCR
                            string levelRaw = QuickOcr(engine, tempPath, LevelRegion);
                            string artifactBonusRaw = QuickOcr(engine, tempPath, ArtifactBONUSRegion);
                            string baseAtkRaw = QuickOcr(engine, tempPath, BaseAtkLabelRegion);
                            string hpCheckRaw = QuickOcr(engine, tempPath, HPRegionCHEAK);
                            string critRateRaw = QuickOcr(engine, tempPath, CritRateRegion);
                            string rechargeRaw = QuickOcr(engine, tempPath, RechargeRegionCHEAK);

                            // 1. Доп.характеристики — проверяем по уже прочитанным строкам, без повторного чтения
                            string hpCheckLow = hpCheckRaw.ToLower();
                            string critRateLow = critRateRaw.ToLower();
                            string rechargeLow = rechargeRaw.ToLower();
                            bool isExtraStats = (hpCheckLow.Contains("hp") || hpCheckLow.Contains("макс"));

                            if (isExtraStats)
                            {
                                Log("  [тип] Доп.характеристики — отложен до 2-го прохода", ConsoleColor.Magenta);
                                extraStatsScreenshots.Add((tempPath, filePath, fileName));
                                continue;
                            }

                            var data = new CharacterData { FileName = fileName };

                            // 2. Оружие
                            string baseAtkLow = baseAtkRaw.ToLower();
                            bool isWeaponScreen = baseAtkLow.Contains("баз") || baseAtkLow.Contains("атак");

                            // 3. Артефакты
                            string hpArtRaw = QuickOcr(engine, tempPath, HPRegion);  // читаем только если нужно
                            bool isArtifactScreen = artifactBonusRaw.Contains("Бонус")
                                                 || hpArtRaw.ToLower().Contains("артефакт");

                            // 4. Таланты — только если не оружие (самая частая ошибочная ветка)
                            bool isTalentScreen = false;
                            if (!isWeaponScreen && !isArtifactScreen)
                            {
                                string t1 = QuickOcr(engine, tempPath, Talent1Region).ToLower();
                                string t2 = QuickOcr(engine, tempPath, Talent2Region).ToLower();
                                string t3 = QuickOcr(engine, tempPath, Talent3Region).ToLower();
                                Log($"  [talent_debug] t1='{t1}' t2='{t2}' t3='{t3}'", ConsoleColor.DarkMagenta);
                                isTalentScreen = t1.Contains("ур") || t2.Contains("ур") || t3.Contains("ур");
                            }

                            if (isWeaponScreen)
                                ProcessWeaponScreenshot(engine, tempPath, data, filePath);
                            else if (isArtifactScreen)
                                ProcessArtifactScreenshot(engine, tempPath, data, filePath);
                            else if (isTalentScreen)
                                ProcessTalentScreenshot(engine, tempPath, data, filePath);
                            else
                                ProcessCharacterScreenshot(engine, tempPath, data, filePath);

                            mainScreenshots.Add(data);
                            File.Delete(tempPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("  Ошибка: " + ex.Message, ConsoleColor.Red);
                    }
                }

                // ── ПРОХОД 2: скриншоты Доп.характеристик ────────────────────────────
                var mergedAfterPass1 = MergeDataByCharacter(mainScreenshots);

                foreach (var (tempPath, origPath, fileName) in extraStatsScreenshots)
                {
                    Log($"\nОбрабатываю (Доп.): {fileName}", ConsoleColor.Cyan);
                    try
                    {
                        var data = new CharacterData { FileName = fileName };
                        ProcessExtraStatsScreenshot(engine, tempPath, data, origPath, mergedAfterPass1);
                        mainScreenshots.Add(data);
                        File.Delete(tempPath);
                    }
                    catch (Exception ex)
                    {
                        Log("  Ошибка: " + ex.Message, ConsoleColor.Red);
                        try { File.Delete(tempPath); } catch { }
                    }
                }
            }

            var mergedResults = MergeDataByCharacter(mainScreenshots);
            PrintResults(mainScreenshots.Count, mergedResults);

            ExportToExcel(mergedResults);
            logWriter.Flush();
            logWriter.Close();
            logFileStream.Close();

            Log("(Лог Tesseract сохранён в tesseract_log.txt)", ConsoleColor.DarkGray);
            Console.ReadLine();
        }
        #endregion

        #region ОПРЕДЕЛЕНИЕ ПАПКИ
        private static void PrepareDebugFolder()
        {
            if (Directory.Exists(DebugPath))
            {
                try
                {
                    foreach (string file in Directory.GetFiles(DebugPath))
                        File.Delete(file);
                    foreach (string dir in Directory.GetDirectories(DebugPath))
                        Directory.Delete(dir, true);
                }
                catch (Exception ex)
                {
                    Log($"Не удалось очистить папку {DebugPath}: {ex.Message}", ConsoleColor.Red);
                }
            }
            else
            {
                Directory.CreateDirectory(DebugPath);
            }
        }
        #endregion

        #region ОБРАБОТКА ФОТО
        private static Bitmap EnhanceImage(Bitmap original)
        {
            var result = new Bitmap(original.Width, original.Height);
            using (var g = Graphics.FromImage(result))
            {
                float contrast = 1.2f;
                float brightness = 0.02f;
                var colorMatrix = new float[][]
                {
                    new float[] {contrast, 0, 0, 0, 0},
                    new float[] {0, contrast, 0, 0, 0},
                    new float[] {0, 0, contrast, 0, 0},
                    new float[] {0, 0, 0, 1, 0},
                    new float[] {brightness, brightness, brightness, 0, 1}
                };

                using (var attrs = new ImageAttributes())
                {
                    attrs.SetColorMatrix(new ColorMatrix(colorMatrix));
                    g.DrawImage(original,
                        new Rectangle(0, 0, original.Width, original.Height),
                        0, 0, original.Width, original.Height,
                        GraphicsUnit.Pixel, attrs);
                }
            }
            return result;
        }
        #endregion

        #region ПРЕДОБРАБОТКА ДЛЯ OCR (ОПТИМИЗИРОВАННАЯ)

        private static string ExtractTextFromRegion(TesseractEngine engine, string imagePath,
            Rectangle region, string regionName,
            bool forTypeDetection = false,
            ElementType? element = null,
            float scale = 6.0f)
        {
            try
            {
                using (var fullImage = new Bitmap(imagePath))
                {
                    var safeRegion = Rectangle.Intersect(region,
                        new Rectangle(0, 0, fullImage.Width, fullImage.Height));
                    if (safeRegion.IsEmpty) return string.Empty;

                    using (var cropped = fullImage.Clone(safeRegion, fullImage.PixelFormat))
                    using (var scaled = ScaleBitmapFast(cropped, scale))
                    {
                        string result = TryOptimalOcr(engine, scaled, element, regionName);

                        if (string.IsNullOrWhiteSpace(result) && element != null)
                        {
                            using (var inverted = InvertBitmapFast(scaled))
                                result = TryOcrWithBitmap(engine, inverted);
                        }

                        if (!forTypeDetection)
                        {
                            string displayText = string.IsNullOrEmpty(result) ? "[пусто]" : result;
                            Log($"  [{regionName}] → {displayText}", ConsoleColor.White);
                        }

                        return result ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                if (!forTypeDetection)
                    Log($"  Ошибка в регионе {regionName}: {ex.Message}", ConsoleColor.Red);
                return string.Empty;
            }
        }

        private static string TryOptimalOcr(TesseractEngine engine, Bitmap scaled,
            ElementType? element, string regionName)
        {
            // Для артефактов - специальная обработка
            if (regionName == "artifact_set" || regionName.Contains("artifact"))
            {
                using (var greenChannel = ExtractGreenChannelFast(scaled))
                using (var binarized = AdaptiveBinarizeFast(greenChannel))
                {
                    string result = TryOcrWithBitmap(engine, binarized);
                    if (!string.IsNullOrWhiteSpace(result)) return result;
                }

                using (var binarized = AdaptiveBinarizeFast(scaled))
                    return TryOcrWithBitmap(engine, binarized);
            }

            var bestChannel = DetermineBestChannel(scaled, element);

            if (bestChannel != null)
            {
                using (var binarized = AdaptiveBinarizeFast(bestChannel))
                {
                    string result = TryOcrWithBitmap(engine, binarized);
                    if (!string.IsNullOrWhiteSpace(result)) return result;
                }
            }

            using (var binarized = AdaptiveBinarizeFast(scaled))
                return TryOcrWithBitmap(engine, binarized);
        }

        private static Bitmap DetermineBestChannel(Bitmap scaled, ElementType? element)
        {
            if (element != null)
            {
                switch (element)
                {
                    case ElementType.Pyro:
                        return ExtractRedChannelFast(scaled);
                    case ElementType.Electro:
                    case ElementType.Geo:
                    case ElementType.Hydro:
                        return ExtractBlueChannelFast(scaled);
                    case ElementType.Cryo:
                        return ExtractBlueChannelFast(scaled);
                    case ElementType.Anemo:
                    case ElementType.Dendro:
                        return ExtractGreenChannelFast(scaled);
                }
            }

            return AnalyzeBestChannelByHistogram(scaled);
        }

        private static Bitmap AnalyzeBestChannelByHistogram(Bitmap source)
        {
            int sampleStep = 4;
            int redIntensity = 0, greenIntensity = 0, blueIntensity = 0;
            int samples = 0;

            var srcData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height),
                ImageLockMode.ReadOnly, source.PixelFormat);

            unsafe
            {
                byte* ptr = (byte*)srcData.Scan0;
                int bytesPerPixel = Image.GetPixelFormatSize(source.PixelFormat) / 8;

                for (int y = 0; y < source.Height; y += sampleStep)
                {
                    byte* row = ptr + y * srcData.Stride;
                    for (int x = 0; x < source.Width; x += sampleStep)
                    {
                        byte b = row[x * bytesPerPixel];
                        byte g = row[x * bytesPerPixel + 1];
                        byte r = row[x * bytesPerPixel + 2];

                        redIntensity += r;
                        greenIntensity += g;
                        blueIntensity += b;
                        samples++;
                    }
                }
            }
            source.UnlockBits(srcData);

            // Выбираем канал с максимальной контрастностью
            if (greenIntensity > redIntensity && greenIntensity > blueIntensity)
                return ExtractGreenChannelFast(source);
            if (redIntensity > blueIntensity)
                return ExtractRedChannelFast(source);

            return ExtractBlueChannelFast(source);
        }

        #region БЫСТРЫЕ ОПЕРАЦИИ С BITMAP (LockBits)

        private static Bitmap ScaleBitmapFast(Bitmap source, float scale)
        {
            int newW = (int)(source.Width * scale);
            int newH = (int)(source.Height * scale);
            var result = new Bitmap(newW, newH);
            using (var g = Graphics.FromImage(result))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.DrawImage(source, 0, 0, newW, newH);
            }
            return result;
        }


        private static Bitmap ExtractRedChannelFast(Bitmap source)
        {
            var result = new Bitmap(source.Width, source.Height, source.PixelFormat);
            var srcData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height),
                ImageLockMode.ReadOnly, source.PixelFormat);
            var dstData = result.LockBits(new Rectangle(0, 0, result.Width, result.Height),
                ImageLockMode.WriteOnly, result.PixelFormat);

            unsafe
            {
                byte* src = (byte*)srcData.Scan0;
                byte* dst = (byte*)dstData.Scan0;
                int bytesPerPixel = Image.GetPixelFormatSize(source.PixelFormat) / 8;

                for (int y = 0; y < source.Height; y++)
                {
                    byte* srcRow = src + y * srcData.Stride;
                    byte* dstRow = dst + y * dstData.Stride;

                    for (int x = 0; x < source.Width; x++)
                    {
                        byte b = srcRow[x * bytesPerPixel];
                        byte g = srcRow[x * bytesPerPixel + 1];
                        byte r = srcRow[x * bytesPerPixel + 2];

                        int redness = r - Math.Max(g, b);
                        byte brightness = redness > 30 ? (byte)(0.299 * r + 0.587 * g + 0.114 * b) : (byte)0;

                        dstRow[x * bytesPerPixel] = brightness;
                        dstRow[x * bytesPerPixel + 1] = brightness;
                        dstRow[x * bytesPerPixel + 2] = brightness;
                        if (bytesPerPixel == 4)
                            dstRow[x * bytesPerPixel + 3] = 255;
                    }
                }
            }

            source.UnlockBits(srcData);
            result.UnlockBits(dstData);
            return result;
        }

        private static Bitmap ExtractGreenChannelFast(Bitmap source)
        {
            var result = new Bitmap(source.Width, source.Height, source.PixelFormat);
            var srcData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height),
                ImageLockMode.ReadOnly, source.PixelFormat);
            var dstData = result.LockBits(new Rectangle(0, 0, result.Width, result.Height),
                ImageLockMode.WriteOnly, result.PixelFormat);

            unsafe
            {
                byte* src = (byte*)srcData.Scan0;
                byte* dst = (byte*)dstData.Scan0;
                int bytesPerPixel = Image.GetPixelFormatSize(source.PixelFormat) / 8;

                for (int y = 0; y < source.Height; y++)
                {
                    byte* srcRow = src + y * srcData.Stride;
                    byte* dstRow = dst + y * dstData.Stride;

                    for (int x = 0; x < source.Width; x++)
                    {
                        byte b = srcRow[x * bytesPerPixel];
                        byte g = srcRow[x * bytesPerPixel + 1];
                        byte r = srcRow[x * bytesPerPixel + 2];

                        int greenness = g - Math.Max(r, b);
                        byte brightness = greenness > 40 ? (byte)(0.299 * r + 0.587 * g + 0.114 * b) : (byte)0;

                        dstRow[x * bytesPerPixel] = brightness;
                        dstRow[x * bytesPerPixel + 1] = brightness;
                        dstRow[x * bytesPerPixel + 2] = brightness;
                        if (bytesPerPixel == 4)
                            dstRow[x * bytesPerPixel + 3] = 255;
                    }
                }
            }

            source.UnlockBits(srcData);
            result.UnlockBits(dstData);
            return result;
        }

        private static Bitmap ExtractBlueChannelFast(Bitmap source)
        {
            var result = new Bitmap(source.Width, source.Height, source.PixelFormat);
            var srcData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height),
                ImageLockMode.ReadOnly, source.PixelFormat);
            var dstData = result.LockBits(new Rectangle(0, 0, result.Width, result.Height),
                ImageLockMode.WriteOnly, result.PixelFormat);

            unsafe
            {
                byte* src = (byte*)srcData.Scan0;
                byte* dst = (byte*)dstData.Scan0;
                int bytesPerPixel = Image.GetPixelFormatSize(source.PixelFormat) / 8;

                for (int y = 0; y < source.Height; y++)
                {
                    byte* srcRow = src + y * srcData.Stride;
                    byte* dstRow = dst + y * dstData.Stride;

                    for (int x = 0; x < source.Width; x++)
                    {
                        byte b = srcRow[x * bytesPerPixel];
                        byte g = srcRow[x * bytesPerPixel + 1];
                        byte r = srcRow[x * bytesPerPixel + 2];

                        int blueness = b - (r + g) / 2;
                        byte brightness = blueness > 80 ? (byte)(0.299 * r + 0.587 * g + 0.114 * b) : (byte)0;

                        dstRow[x * bytesPerPixel] = brightness;
                        dstRow[x * bytesPerPixel + 1] = brightness;
                        dstRow[x * bytesPerPixel + 2] = brightness;
                        if (bytesPerPixel == 4)
                            dstRow[x * bytesPerPixel + 3] = 255;
                    }
                }
            }

            source.UnlockBits(srcData);
            result.UnlockBits(dstData);
            return result;
        }

        private static Bitmap AdaptiveBinarizeFast(Bitmap source)
        {
            // Оптимизация: конвертация в серый + бинаризация за один проход
            var result = new Bitmap(source.Width, source.Height, source.PixelFormat);
            var srcData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height),
                ImageLockMode.ReadOnly, source.PixelFormat);
            var dstData = result.LockBits(new Rectangle(0, 0, result.Width, result.Height),
                ImageLockMode.WriteOnly, result.PixelFormat);

            // Быстрый расчет гистограммы
            int[] histogram = new int[256];
            unsafe
            {
                byte* src = (byte*)srcData.Scan0;
                int bytesPerPixel = Image.GetPixelFormatSize(source.PixelFormat) / 8;

                for (int y = 0; y < source.Height; y++)
                {
                    byte* row = src + y * srcData.Stride;
                    for (int x = 0; x < source.Width; x++)
                    {
                        byte brightness = (byte)(0.299 * row[x * bytesPerPixel + 2] +
                                                 0.587 * row[x * bytesPerPixel + 1] +
                                                 0.114 * row[x * bytesPerPixel]);
                        histogram[brightness]++;
                    }
                }
            }

            // Метод Отсу
            int threshold = CalculateOtsuThresholdFast(histogram, source.Width * source.Height);
            threshold = (int)(threshold * 0.85f);

            // Бинаризация
            unsafe
            {
                byte* src = (byte*)srcData.Scan0;
                byte* dst = (byte*)dstData.Scan0;
                int bytesPerPixel = Image.GetPixelFormatSize(source.PixelFormat) / 8;

                for (int y = 0; y < source.Height; y++)
                {
                    byte* srcRow = src + y * srcData.Stride;
                    byte* dstRow = dst + y * dstData.Stride;

                    for (int x = 0; x < source.Width; x++)
                    {
                        byte brightness = (byte)(0.299 * srcRow[x * bytesPerPixel + 2] +
                                                 0.587 * srcRow[x * bytesPerPixel + 1] +
                                                 0.114 * srcRow[x * bytesPerPixel]);
                        byte color = brightness >= threshold ? (byte)255 : (byte)0;

                        dstRow[x * bytesPerPixel] = color;
                        dstRow[x * bytesPerPixel + 1] = color;
                        dstRow[x * bytesPerPixel + 2] = color;
                        if (bytesPerPixel == 4)
                            dstRow[x * bytesPerPixel + 3] = 255;
                    }
                }
            }

            source.UnlockBits(srcData);
            result.UnlockBits(dstData);
            return result;
        }

        private static int CalculateOtsuThresholdFast(int[] histogram, int totalPixels)
        {
            float sum = 0;
            for (int i = 0; i < 256; i++)
                sum += i * histogram[i];

            float sumB = 0;
            int wB = 0;
            float maxVariance = 0;
            int threshold = 0;

            for (int i = 0; i < 256; i++)
            {
                wB += histogram[i];
                if (wB == 0) continue;

                int wF = totalPixels - wB;
                if (wF == 0) break;

                sumB += i * histogram[i];
                float mB = sumB / wB;
                float mF = (sum - sumB) / wF;

                float variance = wB * wF * (mB - mF) * (mB - mF);
                if (variance > maxVariance)
                {
                    maxVariance = variance;
                    threshold = i;
                }
            }

            return threshold;
        }

        private static Bitmap InvertBitmapFast(Bitmap source)
        {
            var result = new Bitmap(source.Width, source.Height, source.PixelFormat);
            var srcData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height),
                ImageLockMode.ReadOnly, source.PixelFormat);
            var dstData = result.LockBits(new Rectangle(0, 0, result.Width, result.Height),
                ImageLockMode.WriteOnly, result.PixelFormat);

            unsafe
            {
                byte* src = (byte*)srcData.Scan0;
                byte* dst = (byte*)dstData.Scan0;
                int bytesPerPixel = Image.GetPixelFormatSize(source.PixelFormat) / 8;
                int totalBytes = srcData.Stride * source.Height;

                for (int i = 0; i < totalBytes; i += bytesPerPixel)
                {
                    dst[i] = (byte)(255 - src[i]);
                    dst[i + 1] = (byte)(255 - src[i + 1]);
                    dst[i + 2] = (byte)(255 - src[i + 2]);
                    if (bytesPerPixel == 4)
                        dst[i + 3] = src[i + 3];
                }
            }

            source.UnlockBits(srcData);
            result.UnlockBits(dstData);
            return result;
        }

        private static Bitmap PreprocessForNumbersFast(Bitmap source)
        {
            var result = new Bitmap(source.Width, source.Height, source.PixelFormat);
            var srcData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height),
                ImageLockMode.ReadOnly, source.PixelFormat);
            var dstData = result.LockBits(new Rectangle(0, 0, result.Width, result.Height),
                ImageLockMode.WriteOnly, result.PixelFormat);

            // Быстрая гистограмма
            int[] histogram = new int[256];
            unsafe
            {
                byte* src = (byte*)srcData.Scan0;
                int bytesPerPixel = Image.GetPixelFormatSize(source.PixelFormat) / 8;

                for (int y = 0; y < source.Height; y++)
                {
                    byte* row = src + y * srcData.Stride;
                    for (int x = 0; x < source.Width; x++)
                    {
                        byte brightness = (byte)(0.299 * row[x * bytesPerPixel + 2] +
                                                 0.587 * row[x * bytesPerPixel + 1] +
                                                 0.114 * row[x * bytesPerPixel]);
                        histogram[brightness]++;
                    }
                }
            }

            int threshold = CalculateOtsuThresholdFast(histogram, source.Width * source.Height);
            threshold = Math.Max(100, Math.Min(200, threshold));

            // Бинаризация
            unsafe
            {
                byte* src = (byte*)srcData.Scan0;
                byte* dst = (byte*)dstData.Scan0;
                int bytesPerPixel = Image.GetPixelFormatSize(source.PixelFormat) / 8;

                for (int y = 0; y < source.Height; y++)
                {
                    byte* srcRow = src + y * srcData.Stride;
                    byte* dstRow = dst + y * dstData.Stride;

                    for (int x = 0; x < source.Width; x++)
                    {
                        byte brightness = (byte)(0.299 * srcRow[x * bytesPerPixel + 2] +
                                                 0.587 * srcRow[x * bytesPerPixel + 1] +
                                                 0.114 * srcRow[x * bytesPerPixel]);
                        byte color = brightness >= threshold ? (byte)255 : (byte)0;

                        dstRow[x * bytesPerPixel] = color;
                        dstRow[x * bytesPerPixel + 1] = color;
                        dstRow[x * bytesPerPixel + 2] = color;
                        if (bytesPerPixel == 4)
                            dstRow[x * bytesPerPixel + 3] = 255;
                    }
                }
            }

            source.UnlockBits(srcData);
            result.UnlockBits(dstData);
            return result;
        }

        public static ElementType? ExtractElementFromName(string rawName)
        {
            if (string.IsNullOrEmpty(rawName)) return null;

            // Паттерн: "Электро / Синобу" или "Пиро / Ху Тао"
            var slashMatch = Regex.Match(rawName, @"([А-Яа-я]+)\s*/\s*(.+)");
            if (slashMatch.Success)
            {
                string elementRu = slashMatch.Groups[1].Value.Trim();

                switch (elementRu)
                {
                    case "Пиро": return ElementType.Pyro;
                    case "Гидро": return ElementType.Hydro;
                    case "Электро": return ElementType.Electro;
                    case "Крио": return ElementType.Cryo;
                    case "Анемо": return ElementType.Anemo;
                    case "Гео": return ElementType.Geo;
                    case "Дендро": return ElementType.Dendro;
                    default: return null;
                }
            }

            return null;
        }

        // Добавьте этот метод для отладки
        public static void SaveDebugImageForFile(string imagePath, Dictionary<string, Rectangle> regions)
        {
            try
            {
                if (regions == null || regions.Count == 0) return;

                string fileName = Path.GetFileNameWithoutExtension(imagePath);
                string debugFilePath = Path.Combine(DebugPath, $"DEBUG_{fileName}.png");

                using (var fullImage = new Bitmap(imagePath))
                using (var debugImage = new Bitmap(fullImage))
                using (var g = Graphics.FromImage(debugImage))
                {
                    foreach (var region in regions)
                        g.DrawRectangle(new Pen(Color.Red, 7), region.Value);

                    debugImage.Save(debugFilePath, ImageFormat.Png);
                }
            }
            catch (Exception ex) { Log($"  Ошибка сохранения отладочного изображения: {ex.Message}", ConsoleColor.Red); }
        }

        public enum ElementType
        {
            Pyro,    // Красный фон
            Hydro,   // Синий фон
            Electro, // Фиолетовый фон
            Cryo,    // Голубой/Синий фон
            Anemo,   // Бирюзовый/Зелёный фон
            Geo,     // Жёлтый/Оранжевый фон
            Dendro   // Зелёный фон
        }
        #endregion

        private static string QuickOcr(TesseractEngine engine, string imagePath, Rectangle region)
        {
            try
            {
                using (var fullImage = new Bitmap(imagePath))
                {
                    var safeRegion = Rectangle.Intersect(region,
                        new Rectangle(0, 0, fullImage.Width, fullImage.Height));
                    if (safeRegion.IsEmpty) return string.Empty;

                    using (var cropped = fullImage.Clone(safeRegion, fullImage.PixelFormat))
                    using (var processed = PreprocessForNumbersFast(cropped))
                    {
                        return TryOcrWithBitmap(engine, processed);
                    }
                }
            }
            catch { return string.Empty; }
        }

        private static string TryOcrWithBitmap(TesseractEngine engine, Bitmap bmp)
        {
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                using (var pix = Pix.LoadFromMemory(ms.ToArray()))
                using (var page = engine.Process(pix))
                {
                    return page.GetText().Trim().Replace("\n", " ").Replace("\r", "");
                }
            }
        }

        #endregion
        #region ОБРАБОТКА СКРИНШОТА
        private static void ProcessCharacterScreenshot(TesseractEngine engine, string tempPath,
            CharacterData data, string originalImagePath)
        {
            Log("  [тип] Скриншот персонажа", ConsoleColor.Yellow);
            string nameRaw = ExtractTextFromRegion(engine, tempPath, NameRegion, "name");
            ElementType? element = ExtractElementFromName(nameRaw);
            Log($"  [элемент] {element?.ToString() ?? "неизвестен"}", ConsoleColor.Cyan);
            string levelRaw = ExtractTextFromRegion(engine, tempPath, LevelRegion, "level", element: element);
            string hpRaw = ExtractTextFromRegion(engine, tempPath, HPRegion, "hp", element: element);
            string atkRaw = ExtractTextFromRegion(engine, tempPath, AtkRegion, "atk", element: element);
            string defRaw = ExtractTextFromRegion(engine, tempPath, DefRegion, "def", element: element);
            string msRaw = ExtractTextFromRegion(engine, tempPath, MsRegion, "ms", element: element);

            if (!string.IsNullOrEmpty(nameRaw)) data.CharacterName = ExtractCharacterName(nameRaw);
            if (!string.IsNullOrEmpty(levelRaw)) data.CharacterLevel = CleanLevel(levelRaw);

            data.HP = CleanStat(hpRaw);
            data.Atk = CleanStat(atkRaw);
            data.Def = CleanStat(defRaw);
            data.MasteryStihiy = CleanStat(msRaw);

            var regions = new Dictionary<string, Rectangle>
            {
                { "name",  NameRegion  },
                { "level", LevelRegion },
                { "hp",    HPRegion    },
                { "atk",   AtkRegion   },
                { "def",   DefRegion   },
                { "ms",    MsRegion    },
            };
            SaveDebugImageForFile(originalImagePath, regions);
        }

        private static void ProcessWeaponScreenshot(TesseractEngine engine, string tempPath,
            CharacterData data, string originalImagePath)
        {
            Log("  [тип] Скриншот оружия", ConsoleColor.Yellow);
            string nameRaw = ExtractTextFromRegion(engine, tempPath, NameRegion, "name");
            ElementType? element = ExtractElementFromName(nameRaw);
            Log($"  [элемент] {element?.ToString() ?? "неизвестен"}", ConsoleColor.Cyan);
            string weaponNameRaw = ExtractTextFromRegion(engine, tempPath, NameWeaponRegion, "weapon", element: element);
            string weaponLevelRaw = ExtractTextFromRegion(engine, tempPath, LevelWeaponRegion, "weapon_level", element: element);

            if (!string.IsNullOrEmpty(nameRaw)) data.CharacterName = ExtractCharacterName(nameRaw);

            if (!string.IsNullOrEmpty(weaponNameRaw))
                if (string.IsNullOrEmpty(data.CharacterName) || !weaponNameRaw.Contains(data.CharacterName))
                    data.WeaponName = CleanWeaponName(weaponNameRaw); //ApplyWeaponCorrections(CleanWeaponName(weaponNameRaw));

            if (!string.IsNullOrEmpty(weaponLevelRaw)) data.WeaponLevel = CleanLevel(weaponLevelRaw);

            var regions = new Dictionary<string, Rectangle>
            {
                { "name",         NameRegion        },
                { "weapon",       NameWeaponRegion   },
                { "weapon_level", LevelWeaponRegion  }
            };
            SaveDebugImageForFile(originalImagePath, regions);
        }

        private static void ProcessArtifactScreenshot(TesseractEngine engine, string tempPath,
            CharacterData data, string originalImagePath)
        {
            Log("  [тип] Скриншот артефактов", ConsoleColor.Yellow);

            string nameRaw = ExtractTextFromRegion(engine, tempPath, NameRegion, "name");
            ElementType? element = ExtractElementFromName(nameRaw);
            Log($"  [элемент] {element?.ToString() ?? "неизвестен"}", ConsoleColor.Cyan);
            string artifactRaw = ExtractTextFromRegion(engine, tempPath, ArtifactSetRegion, "artifact_set", element: element);

            if (!string.IsNullOrEmpty(nameRaw)) data.CharacterName = ExtractCharacterName(nameRaw);
            if (!string.IsNullOrEmpty(artifactRaw)) data.ArtifactSet = ParseArtifactSet(artifactRaw);

            var regions = new Dictionary<string, Rectangle>
            {
                { "name",         NameRegion        },
                { "artifact_set", ArtifactSetRegion }
            };
            SaveDebugImageForFile(originalImagePath, regions);
        }

        private static void ProcessTalentScreenshot(TesseractEngine engine, string tempPath,
            CharacterData data, string originalImagePath)
        {
            Log("  [тип] Скриншот талантов", ConsoleColor.Yellow);

            string nameRaw = ExtractTextFromRegion(engine, tempPath, NameRegion, "name");
            ElementType? element = ExtractElementFromName(nameRaw);
            Log($"  [элемент] {element?.ToString() ?? "неизвестен"}", ConsoleColor.Cyan);
            string t1Raw = ExtractTextFromRegion(engine, tempPath, Talent1Region, "talent1", element: element);
            string t2Raw = ExtractTextFromRegion(engine, tempPath, Talent2Region, "talent2", element: element);
            string t3Raw = ExtractTextFromRegion(engine, tempPath, Talent3Region, "talent_3", element: element);

            if (!string.IsNullOrEmpty(nameRaw)) data.CharacterName = ExtractCharacterName(nameRaw);

            int T2bonus = IsBlueTalent(tempPath, Talent2Region) ? 3 : 0;
            int T3bonus = IsBlueTalent(tempPath, Talent3Region) ? 3 : 0;

            data.Talent1 = ParseTalentLevel(t1Raw);
            data.Talent2 = (int.Parse(ParseTalentLevel(t2Raw)) - T2bonus).ToString();
            data.Talent3 = (int.Parse(ParseTalentLevel(t3Raw)) - T3bonus).ToString();
            Log($"{data.Talent1}, {data.Talent2}, {data.Talent3}");
            var regions = new Dictionary<string, Rectangle>
            {
                { "name",        NameRegion       },
                { "talent1",     Talent1Region    },
                { "talent2",     Talent2Region    },
                { "talent_3",  Talent3Region  },
            };
            SaveDebugImageForFile(originalImagePath, regions);
        }
        private static bool IsBlueTalent(string imagePath, Rectangle region)
        {
            try
            {
                using (var image = new Bitmap(imagePath))
                using (var cropped = image.Clone(region, image.PixelFormat))
                {
                    int cyanCount = 0;
                    int total = 0;

                    for (int y = 0; y < cropped.Height; y += 2)
                    {
                        for (int x = 0; x < cropped.Width; x += 2)
                        {
                            Color p = cropped.GetPixel(x, y);

                            // Ключевое отличие голубого текста: G и B значительно выше R
                            int gbAvg = (p.G + p.B) / 2;
                            if (gbAvg - p.R > 100 && gbAvg > 150)
                                cyanCount++;

                            total++;
                        }
                    }

                    return (float)cyanCount / total > 0.008f;
                }
            }
            catch
            {
                return false;
            }
        }

        private static void ProcessExtraStatsScreenshot(TesseractEngine engine, string tempPath,
            CharacterData data, string originalImagePath,
            Dictionary<string, CharacterData> knownCharacters)
        {
            Log("  [тип] Доп.характеристики", ConsoleColor.Magenta);
            string hpText = ExtractTextFromRegion(engine, tempPath, HPRegionCHEAK, "chk_hp");
            string atkText = ExtractTextFromRegion(engine, tempPath, AtkRegionCHEAK, "chk_atk");
            string defText = ExtractTextFromRegion(engine, tempPath, DefRegionCHEAK, "chk_def");
            string msText = ExtractTextFromRegion(engine, tempPath, MsRegionCHEAK, "chk_ms");
            string critRateText = ExtractTextFromRegion(engine, tempPath, CritRateRegion, "crit_rate");
            string critDmgText = ExtractTextFromRegion(engine, tempPath, CritDMGRegion, "crit_dmg");
            string rechargeText = ExtractTextFromRegion(engine, tempPath, RechargeRegionCHEAK, "recharge");

            string hpVal = CleanStatFromLabel(hpText, isPercent: false);
            string atkVal = CleanStatFromLabel(atkText, isPercent: false);
            string defVal = CleanStatFromLabel(defText, isPercent: false);
            string msVal = CleanStatFromLabel(msText, isPercent: false);
            string critRate = CleanStatFromLabel(critRateText, isPercent: true);
            string critDmg = CleanStatFromLabel(critDmgText, isPercent: true);
            string recharge = CleanStatFromLabel(rechargeText, isPercent: true);

            string matchedName = MatchCharacterByStats(hpVal, atkVal, defVal, msVal, knownCharacters);

            if (matchedName != null)
            {
                Log($"  [матч] Персонаж определён как: '{matchedName}'", ConsoleColor.Green);
                data.CharacterName = matchedName;
            }
            else
            {
                Log("  [!] Персонаж не найден по характеристикам — записываю без имени", ConsoleColor.DarkYellow);
                data.CharacterName = null;
            }

            data.CritRate = critRate;
            data.CritDMG = critDmg;
            data.EnergyRecharge = recharge;
            data.HP = hpVal;
            data.Atk = atkVal;
            data.Def = defVal;
            if (!string.IsNullOrEmpty(msVal)) data.MasteryStihiy = msVal;
            Log($"{data.Talent1}, {data.Talent2}, {data.Talent3}");
            var regions = new Dictionary<string, Rectangle>
            {
                { "chk_hp",   HPRegionCHEAK       },
                { "chk_atk",  AtkRegionCHEAK      },
                { "chk_def",  DefRegionCHEAK      },
                { "chk_ms",   MsRegionCHEAK       },
                { "crit_rate", CritRateRegion     },
                { "crit_dmg",  CritDMGRegion      },
                { "recharge",  RechargeRegionCHEAK },
            };
            SaveDebugImageForFile(originalImagePath, regions);
        }
        #endregion

        #region EXTRA_STATS_HELPERS
        private static bool ValidateExtraStatsLabels(string hpText, string critRateText, string rechargeText)
        {
            string hp = hpText.ToLower();
            string crit = critRateText.ToLower();
            string rech = rechargeText.ToLower();

            bool hasHp = hp.Contains("hp") || hp.Contains("макс");
            bool hasCrit = crit.Contains("крит");
            bool hasRecharge = rech.Contains("восст") || rech.Contains("энерг");

            return hasHp && hasCrit && hasRecharge;
        }

        private static string CleanStatFromLabel(string text, bool isPercent)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            if (isPercent)
            {
                string digits = Regex.Replace(text, @"\D", "");
                if (string.IsNullOrEmpty(digits)) return string.Empty;
                if (decimal.TryParse(digits, out decimal value))
                    return (value / 10).ToString("0.0", System.Globalization.CultureInfo.GetCultureInfo("ru-RU"));
                return string.Empty;
            }
            else
            {
                // Склеиваем пробелы внутри целых чисел: "24 900" → "24900"
                string collapsed = Regex.Replace(text, @"(\d)\s+(\d)", "$1$2");

                if (collapsed.Contains('+'))
                {
                    var numbers = Regex.Matches(collapsed, @"-?\d+")
                        .Cast<Match>()
                        .Select(m => decimal.Parse(m.Value))
                        .ToList();

                    if (numbers.Count > 0)
                    {
                        decimal sum = numbers.Sum();
                        return sum.ToString();
                    }
                }
                var matches = Regex.Matches(collapsed, @"\d+");
                if (matches.Count > 0)
                    return matches[matches.Count - 1].Value;

                return string.Empty;
            }
        }

        private static string MatchCharacterByStats(
            string hp, string atk, string def, string ms,
            Dictionary<string, CharacterData> knownCharacters)
        {
            if (string.IsNullOrEmpty(hp) && string.IsNullOrEmpty(atk) && string.IsNullOrEmpty(def)) return null;

            string bestMatch = null;
            int bestScore = 0;

            foreach (var kvp in knownCharacters)
            {
                var known = kvp.Value;
                int score = 0;

                if (NumericMatch(hp, known.HP)) score++;
                if (NumericMatch(atk, known.Atk)) score++;
                if (NumericMatch(def, known.Def)) score++;
                if (NumericMatch(ms, known.MasteryStihiy)) score++;

                if (score <= bestScore) continue;
                bestScore = score;
                bestMatch = kvp.Key;
            }

            if (bestScore == 0) return null;
            bool hpIsSpecific = !string.IsNullOrEmpty(hp) && hp.Length >= 4;
            int threshold = hpIsSpecific && NumericMatch(hp,
                knownCharacters.TryGetValue(bestMatch ?? "", out var bm) ? bm.HP : null) ? 1 : 2;
            return bestScore >= threshold ? bestMatch : null;
        }

        private static bool NumericMatch(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
            if (!double.TryParse(a, out double va) || !double.TryParse(b, out double vb)) return false;
            if (va == 0 && vb == 0) return false; // оба нуля — не считаем совпадением
            double maxVal = Math.Max(Math.Abs(va), Math.Abs(vb));
            return Math.Abs(va - vb) / maxVal <= 0.05;
        }

        #endregion

        #region ОЧИСТКА ТЕКСТА
        private static string CleanName(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            text = text.TrimStart('\'', '"', ' ', '°', '•', '-', '`', '´');
            text = text.TrimEnd('\'', '"', ' ', '°', '•', '-', '`', '´', ':', '.', '+');
            text = Regex.Replace(text, @"\s+", " ").Trim();
            return text;
        }

        private static string ExtractCharacterName(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return raw;

            string normalized = Regex.Replace(raw, @"(\w)\s+[74|I1Гг/\\]\s+(\w)", "$1 / $2");
            int slashIdx = normalized.IndexOf('/');

            string result = slashIdx >= 0 && slashIdx < normalized.Length - 1
                ? normalized.Substring(slashIdx + 1).Trim()
                : normalized.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries) is var words && words.Length >= 2
                    ? string.Join(" ", words.Skip(1))
                    : raw;

            return string.Join(" ", result.Split(' ').Select(w => char.ToUpper(w[0]) + w.Substring(1).ToLower()));
        }

        private static string CleanWeaponName(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            int start = 0;
            while (start < text.Length && !char.IsLetter(text[start])) start++;
            if (start >= text.Length) return string.Empty;
            text = text.Substring(start);

            var spaceMatch = Regex.Match(text, @"\s{2,}");
            if (spaceMatch.Success) text = text.Substring(0, spaceMatch.Index);

            var punctMatch = Regex.Match(text, @"[.,;!?*_+=/\\|]");
            if (punctMatch.Success && punctMatch.Index > 0) text = text.Substring(0, punctMatch.Index);

            text = Regex.Replace(text, @"[^\p{L}\s]+$", "").Trim();
            text = Regex.Replace(text, @"\s+", " ");

            if (text.Length > 0) text = char.ToUpper(text[0]) + text.Substring(1);
            var words = text.Split(' ');
            if (words.Length == 1 && words[0].Length < 3) return string.Empty;

            return text;
        }

        private static string CleanLevel(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            var slashMatch = Regex.Match(text, @"(\d{1,3})\s*/\s*\d{1,3}");
            if (slashMatch.Success) return slashMatch.Groups[1].Value;
            var match = Regex.Match(text, @"\d+");
            return match.Success ? match.Value : string.Empty;
        }

        private static string CleanStat(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            // Склеиваем пробелы внутри чисел: "2 682" → "2682"
            text = Regex.Replace(text, @"(\d)\s+(\d)", "$1$2");
            // Сначала пробуем найти число после нецифрового префикса (метки)
            var match = Regex.Match(text, @"\D+(\d+)");
            if (match.Success) return match.Groups[1].Value;
            // Fallback: берём последнее число в строке (если метка не распозналась)
            var allNums = Regex.Matches(text, @"\d+");
            return allNums.Count > 0 ? allNums[allNums.Count - 1].Value : "0";
        }

        private static string ParseArtifactSet(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            text = Regex.Replace(text, @"\s+", " ").Trim();

            // Убираем одиночные буквы-мусор (иконки галочек) перед словами
            text = Regex.Replace(text, @"(?<!\w)[А-ЯЁA-Z]\s+(?=[А-ЯЁA-Z])", "");
            text = Regex.Replace(text, @"(?<!\w)(?!в)[а-яёa-z]\s+(?=[А-ЯЁA-Z])", "");
            text = Regex.Replace(text, @"(?<!\w)[А-ЯЁа-яёA-Za-z]\s+(?=2\s*предм)", "");
            text = Regex.Replace(text, @"(?<!\w)[А-ЯЁA-Z]{2,3}\s+(?=[А-ЯЁA-Z])", "");
            text = Regex.Replace(text, @"[А-ЯЁа-яёA-Za-z](\w)\1{2,}", "");

            if (text.Contains("нет")) return "-";

            if (Regex.IsMatch(text, @"4\s*предм"))
            {
                var match = Regex.Match(text, @"^(.+?)\s*2\s*предм");
                if (match.Success) return CleanName(match.Groups[1].Value) + " (4 шт.)";
            }

            var twoSetMatches = Regex.Matches(text, @"([А-ЯЁа-яёA-Za-z][^2]+?)\s*2\s*предм");
            if (twoSetMatches.Count >= 2)
                return $"{CleanName(twoSetMatches[0].Groups[1].Value)} (2 шт.) + {CleanName(twoSetMatches[1].Groups[1].Value)} (2 шт.)";
            if (twoSetMatches.Count == 1)
                return CleanName(twoSetMatches[0].Groups[1].Value) + " (2 шт.)";

            return string.Empty;
        }

        private static string ParseTalentLevel(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            // Нормализуем OCR-замены цифр буквами
            text = NormalizeTalentOcrErrors(text);
            text = Regex.Replace(text, @"([Уу][Рр])\s*[гГ]\s*(?=\d)", "$1 ");
            var match = Regex.Match(text, @"[Уу][Рр]\.?\s*(\d+)");
            return match.Success ? match.Groups[1].Value : string.Empty;
        }

        private static string ParseTalentLevelMax(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            text = NormalizeTalentOcrErrors(text);

            var matches = Regex.Matches(text, @"[Уу][Рр]\.?\s*(\d+)");
            if (matches.Count == 0) return string.Empty;
            int max = 0;
            foreach (Match m in matches)
                if (int.TryParse(m.Groups[1].Value, out int v) && v > max)
                    max = v;
            return max > 0 ? max.ToString() : string.Empty;
        }

        private static string NormalizeTalentOcrErrors(string text)
        {
            // После "ур." ожидаем цифру — заменяем типичные OCR-ошибки
            return Regex.Replace(text,
                @"([Уу][Рр]\.?\s*)([рРбБлЛзЗоОlIiа])",
                m => m.Groups[1].Value + m.Groups[2].Value
                    .Replace("р", "6").Replace("Р", "6")
                    .Replace("б", "6").Replace("Б", "6")
                    .Replace("л", "1").Replace("Л", "1")
                    .Replace("з", "3").Replace("З", "3")
                    .Replace("о", "0").Replace("О", "0")
                    .Replace("l", "1").Replace("I", "1").Replace("i", "1")
                    .Replace("а", "4")
            );
        }
        #endregion

        #region DATA_MERGING
        private static Dictionary<string, CharacterData> MergeDataByCharacter(List<CharacterData> allScreenshots)
        {
            var merged = new Dictionary<string, CharacterData>();

            foreach (var screenshot in allScreenshots)
            {
                if (string.IsNullOrEmpty(screenshot.CharacterName)) continue;

                if (!merged.TryGetValue(screenshot.CharacterName, out var existing))
                {
                    merged[screenshot.CharacterName] = new CharacterData
                    {
                        CharacterName = screenshot.CharacterName,
                        CharacterLevel = screenshot.CharacterLevel,
                        WeaponName = screenshot.WeaponName,
                        WeaponLevel = screenshot.WeaponLevel,
                        HP = screenshot.HP,
                        Atk = screenshot.Atk,
                        Def = screenshot.Def,
                        MasteryStihiy = screenshot.MasteryStihiy,
                        ArtifactSet = string.IsNullOrEmpty(screenshot.ArtifactSet) ? "-" : screenshot.ArtifactSet,
                        CritRate = screenshot.CritRate,
                        CritDMG = screenshot.CritDMG,
                        EnergyRecharge = screenshot.EnergyRecharge,
                        Talent1 = screenshot.Talent1,
                        Talent2 = screenshot.Talent2,
                        Talent3 = screenshot.Talent3,
                    };
                }
                else
                {
                    existing.CharacterLevel = screenshot.CharacterLevel ?? existing.CharacterLevel;
                    existing.WeaponName = screenshot.WeaponName ?? existing.WeaponName;
                    existing.WeaponLevel = screenshot.WeaponLevel ?? existing.WeaponLevel;
                    existing.HP = screenshot.HP ?? existing.HP;
                    existing.Atk = screenshot.Atk ?? existing.Atk;
                    existing.Def = screenshot.Def ?? existing.Def;
                    existing.MasteryStihiy = screenshot.MasteryStihiy ?? existing.MasteryStihiy;
                    if (!string.IsNullOrEmpty(screenshot.ArtifactSet)) existing.ArtifactSet = screenshot.ArtifactSet;
                    else if (string.IsNullOrEmpty(existing.ArtifactSet)) existing.ArtifactSet = "-";
                    existing.CritRate = screenshot.CritRate ?? existing.CritRate;
                    existing.CritDMG = screenshot.CritDMG ?? existing.CritDMG;
                    existing.EnergyRecharge = screenshot.EnergyRecharge ?? existing.EnergyRecharge;
                    existing.Talent1 = screenshot.Talent1 ?? existing.Talent1;
                    existing.Talent2 = screenshot.Talent2 ?? existing.Talent2;
                    existing.Talent3 = screenshot.Talent3 ?? existing.Talent3;

                }
            }
            return merged;
        }
        #endregion

        #region ВЫВОД
        private static void PrintResults(int totalProcessed, Dictionary<string, CharacterData> merged)
        {
            Log("\n=======================================", ConsoleColor.DarkCyan);
            Log($"Обработано скриншотов: {totalProcessed}", ConsoleColor.DarkCyan);
            Log($"Найдено персонажей: {merged.Count}", ConsoleColor.DarkCyan);
            Log("=======================================\n", ConsoleColor.DarkCyan);

            foreach (var character in merged.Values)
            {
                Log($"Персонаж: '{character.CharacterName}'", ConsoleColor.Green);
                string output = null;
                output += $"   Уровень: '{character.CharacterLevel ?? "-"}'";
                output += $"\n   Оружие: '{character.WeaponName ?? "-"}'";
                output += $" {character.WeaponLevel ?? "-"} уровня";
                output += $"\n   HP: {character.HP ?? "-"}";
                output += $"\n   АТК: {character.Atk ?? "-"}";
                output += $"\n   Защита: {character.Def ?? "0"}";
                output += $"\n   МС: {character.MasteryStihiy ?? "0"}";
                output += $"\n   Арт: {character.ArtifactSet ?? "-"}";
                if (!string.IsNullOrEmpty(character.CritRate))
                    output += $"\n   Крит.шанс: {character.CritRate}%";
                if (!string.IsNullOrEmpty(character.CritDMG))
                    output += $"\n   Крит.урон: {character.CritDMG}%";
                if (!string.IsNullOrEmpty(character.EnergyRecharge))
                    output += $"\n   Восст.энергии: {character.EnergyRecharge}%";
                output += $"\n   Т1: {character.Talent1 ?? "-"}";
                output += $"\n   Т2: {character.Talent2 ?? "-"}";
                output += $"\n   Т3: {character.Talent3 ?? "-"}";
                Log(output, ConsoleColor.White);
            }
        }
        #endregion

        #region EXCEL
        private static void ExportToExcel(Dictionary<string, CharacterData> mergedResults)
        {
            if (!File.Exists(ExcelFilePath))
            {
                Log($"Файл Excel не найден: {ExcelFilePath}", ConsoleColor.Red);
                return;
            }

            try
            {
                using (var workbook = new XLWorkbook(ExcelFilePath))
                {
                    var sheet = workbook.Worksheet("Прокачка");

                    int updated = 0;
                    int notFound = 0;

                    foreach (var character in mergedResults.Values)
                    {
                        int row = -1;
                        for (int r = 1; r <= sheet.LastRowUsed().RowNumber(); r++)
                        {
                            var val = sheet.Cell(r, "A").GetString().Trim();
                            if (string.Equals(val, character.CharacterName.Trim(), StringComparison.OrdinalIgnoreCase))
                            {
                                row = r;
                                break;
                            }
                        }

                        if (row == -1)
                        {
                            Log($"  [excel] Персонаж не найден в таблице: '{character.CharacterName}'", ConsoleColor.DarkYellow);
                            notFound++;
                            continue;
                        }

                        bool hasChanges = false;

                        if (!string.IsNullOrEmpty(character.CharacterLevel) && sheet.Cell(row, "B").GetString() != character.CharacterLevel)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.WeaponName) && sheet.Cell(row, "Q").GetString().Trim() != character.WeaponName)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.WeaponLevel) && sheet.Cell(row, "C").GetString() != character.WeaponLevel)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.HP) && sheet.Cell(row, "I").GetString() != character.HP)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.Atk) && sheet.Cell(row, "J").GetString() != character.Atk)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.Def) && sheet.Cell(row, "K").GetString() != character.Def)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.MasteryStihiy) && sheet.Cell(row, "L").GetString() != character.MasteryStihiy)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.ArtifactSet) && sheet.Cell(row, "P").GetString().Trim() != character.ArtifactSet)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.CritRate) && sheet.Cell(row, "N").GetString() != character.CritRate)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.CritDMG) && sheet.Cell(row, "O").GetString() != character.CritDMG)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.EnergyRecharge) && sheet.Cell(row, "M").GetString() != character.EnergyRecharge)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.Talent1) && sheet.Cell(row, "D").GetString() != character.Talent1)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.Talent2) && sheet.Cell(row, "E").GetString() != character.Talent2)
                            hasChanges = true;
                        if (!string.IsNullOrEmpty(character.Talent3) && sheet.Cell(row, "F").GetString() != character.Talent3)
                            hasChanges = true;

                        if (!hasChanges)
                        {
                            Log($"  [excel] Данные повторяются: '{character.CharacterName}'", ConsoleColor.DarkGray);
                            continue;
                        }

                        if (!string.IsNullOrEmpty(character.CharacterLevel))
                            sheet.Cell(row, "B").Value = int.Parse(character.CharacterLevel);
                        if (!string.IsNullOrEmpty(character.WeaponName))
                            sheet.Cell(row, "Q").Value = character.WeaponName;
                        if (!string.IsNullOrEmpty(character.WeaponLevel))
                            sheet.Cell(row, "C").Value = int.Parse(character.WeaponLevel);
                        if (!string.IsNullOrEmpty(character.HP))
                            sheet.Cell(row, "I").Value = int.Parse(character.HP);
                        if (!string.IsNullOrEmpty(character.Atk))
                            sheet.Cell(row, "J").Value = int.Parse(character.Atk);
                        if (!string.IsNullOrEmpty(character.Def))
                            sheet.Cell(row, "K").Value = int.Parse(character.Def);
                        sheet.Cell(row, "L").Value = int.TryParse(character.MasteryStihiy, out int value) ? value : 0;
                        if (!string.IsNullOrEmpty(character.ArtifactSet))
                            sheet.Cell(row, "P").Value = character.ArtifactSet;
                        if (!string.IsNullOrEmpty(character.CritRate))
                            sheet.Cell(row, "N").Value = double.Parse(character.CritRate) / 100;
                        if (!string.IsNullOrEmpty(character.CritDMG))
                            sheet.Cell(row, "O").Value = double.Parse(character.CritDMG) / 100;
                        if (!string.IsNullOrEmpty(character.EnergyRecharge))
                            sheet.Cell(row, "M").Value = double.Parse(character.EnergyRecharge) / 100;
                        if (!string.IsNullOrEmpty(character.Talent1))
                            sheet.Cell(row, "D").Value = int.Parse(character.Talent1);
                        if (!string.IsNullOrEmpty(character.Talent2))
                            sheet.Cell(row, "E").Value = int.Parse(character.Talent2);
                        if (!string.IsNullOrEmpty(character.Talent3))
                            sheet.Cell(row, "F").Value = int.Parse(character.Talent3);

                        Log($"  [excel] Обновлён: '{character.CharacterName}'", ConsoleColor.DarkGreen);
                        updated++;
                    }

                    workbook.Save();

                    Log($"\n  Обновлено персонажей: {updated}", ConsoleColor.Green);
                    if (notFound > 0)
                        Log($"  Не найдено в таблице: {notFound}", ConsoleColor.DarkYellow);
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка при работе с Excel: {ex.Message}", ConsoleColor.Red);
            }
        }
        #endregion
    }
    #endregion
}
