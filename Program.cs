using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using iTextSharp.text.pdf;
#if DEBUG
using System.Diagnostics;
using Serilog;
#endif

namespace GenUuid;

public static class ExtractUuid
{
  private static readonly char[] TRIM_CHAR = new char[] { '{', '[', ']', '}' };
  /* Еще можно добавить попытку извлечения для других форматов МС Офиса,
   * но они либо его в готовом виде не содержат, либо встречаются крайне редко
   * (у меня).
   * Также можно добавить в будущем поддержку XPS, но по нему я дополнительно
   * информации найти не смог :(. */
  private static Dictionary<string, Func<Stream, Guid>> _opFuncDic = new(){
    {".pdf", Pdf},
    {".epub", Epub},
    {".fb2", Fb2},
    {".fbd", Fb2},
    {".docx", Docx}};

  /// <summary>
  /// Извлечение UUID из подходящей строки
  /// </summary>
  /// <param name="s">Строка</param>
  /// <returns>Результат</returns>
  private static Guid GetUuidFromString(this string s) =>
      Guid.Parse(s.Substring(s.LastIndexOf(':') + 1));

  /// <summary>
  /// Изменяет порядок байтов массива в MD5
  /// </summary>
  /// <param name="hash">Массив байтов MD5</param>
  /// <returns>Упорядоченный массив</returns>
  private static byte[] ChangeByteOrder(byte[] hash)
  {
    if (hash == null || hash.Length != 16)
      throw new FormatException("Чё?");
    //HACK: Подозреваю, что могут быть какие-то проблемы на процессорах
    //      с другим порядком байтов
    byte[] sum =
    {
      hash[3], hash[2], hash[1], hash[0], hash[5], hash[4], hash[7], hash[6],
      hash[8], hash[9], hash[10], hash[11], hash[12], hash[13], hash[14], hash[15]
    };
    return sum;
  }

  /// <summary>
  /// Работа с XML в ZIP определённым способом
  /// </summary>
  /// <param name="file">Имя ZIP-файла</param>
  /// <param name="fileInArchive">Файл в архиве</param>
  /// <param name="f">Способ работы</param>
  /// <typeparam name="T">Тип результата</typeparam>
  /// <returns>Результат</returns>
  private static T XmlInZip<T>(Stream stream, string fileInArchive, Func<StreamReader, T> f)
  {
    ZipArchive? arc = default;
    T? result = default;
    try
    {
      arc = new ZipArchive(stream, ZipArchiveMode.Read);
      using var sr = new StreamReader(arc.GetEntry(fileInArchive).Open());
      result = f.Invoke(sr);
    }
    catch (Exception e)
    {
#if DEBUG
      Log.Warning(e.Message);
#endif
    }
    return (T)result;
  }

  /// <summary>
  /// Извлечение UUID из файлов MS Word 2013+
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Результат</returns>
  private static Guid? Docx15(StreamReader stream)
  {
    var xdoc = XDocument.Load(stream);
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    XNamespace w15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    var eid = xdoc?.Element(w + "settings")
        ?.Element(w15 + "docId")
        ?.Attribute(w15 + "val")
        ?.Value;
#if DEBUG
    if (!string.IsNullOrWhiteSpace(eid))
      Log.Information($"DOCX docId.w15 <<{eid}>>");
#endif
    if(Guid.TryParse(eid, out Guid result))
      return result;
    else
      return null;
  }

  /// <summary>
  /// Извлечение нестандартного UUID из Docx
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Результат</returns>
  private static Guid? DocxNS(StreamReader stream)
  {
    var xdoc = XDocument.Load(stream);
    XNamespace od = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    var eid = xdoc?.Element(od + "Properties")
        ?.Elements(od + "property")
        ?.First()
        ?.Attribute("fmtid")
        ?.Value;
#if DEBUG
    if (!string.IsNullOrWhiteSpace(eid))
      Log.Information($"DOCX Property fmtid <<{eid}>>");
#endif
    if(Guid.TryParse(eid, out Guid result))
      return result;
    else
      return null;
  }

  /// <summary>
  /// Находим путь к файлу с информацией в EPUB. Если его нет бросаем исключение
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Путь к файлу .opf</returns>
  private static string Container(StreamReader stream)
  {
    var xdoc = XDocument.Load(stream);
    XNamespace container = "urn:oasis:names:tc:opendocument:xmlns:container";
    var result = xdoc?.Element(container + "container")
        ?.Elements(container + "rootfiles")
        ?.First()
        ?.Element(container + "rootfile")
        ?.Attribute("full-path")
        ?.Value;
#if DEBUG
    if (!string.IsNullOrWhiteSpace(result))
      Log.Information($"EPUB content.opf adress <<{result}>>");
#endif
    if (result == null)
      throw new NullReferenceException();
    else
      return result;
  }

  /// <summary>
  /// Находим UUID в EPUB
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>UUID</returns>
  private static Guid? Content(StreamReader stream)
  {
    var xdoc = XDocument.Load(stream);
    XNamespace opf = "http://www.idpf.org/2007/opf";
    XNamespace dc = "http://purl.org/dc/elements/1.1/";
    string? preresult = null;
    Guid? result = null;

    try
    {
      preresult = xdoc?.Element(opf + "package")
        ?.Element(opf + "metadata")
        ?.Elements(dc + "identifier")
        ?.First(
              x =>
                  (
                    x.Attribute(opf + "scheme")?.Value?.ToLowerInvariant() == "uuid"
                    || x.Attribute(opf + "scheme")?.Value?.ToLowerInvariant() == "calibre"
                  )
          )
          ?.Value;
      if (!string.IsNullOrWhiteSpace(preresult))
      {
#if DEBUG
        Log.Information($"EPUB uuid <<{preresult}>>");
#endif
        result = preresult.GetUuidFromString();
      }
    }
    catch (Exception e)
    {
#if DEBUG
      Log.Warning(e.Message);
#endif
    }

    if (result == null)
    {
      try
      {
        preresult = xdoc?.Element(opf + "package")
            ?.Element(opf + "metadata")
            ?.Elements(dc + "identifier")
            ?.First()
            .Attribute("id")
            ?.Value;
        if (!string.IsNullOrWhiteSpace(preresult))
        {
#if DEBUG
          Log.Information($"EPUB uuid <<{preresult}>>");
#endif
          result = preresult?.GetUuidFromString();
        }
      }
      catch (Exception e)
      {
#if DEBUG
        Log.Warning(e.Message);
#endif
      }
    }

    if (result == null)
    {
      try
      {
        preresult = xdoc?.Element(opf + "package")
        ?.Element(opf + "metadata")
        ?.Elements(dc + "identifier")
        ?.First(x => x.Attribute("id")?.Value?.ToLowerInvariant() == "uuid")
        ?.Value;
        if (!string.IsNullOrWhiteSpace(preresult))
        {
#if DEBUG
          Log.Information($"EPUB uuid <<{preresult}>>");
#endif
          result = preresult.GetUuidFromString();
        }
      }
      catch (Exception e)
      {
#if DEBUG
        Log.Warning(e.Message);
#endif
      }
    }

    if (result == null)
    {
      preresult = xdoc?.Root?.Attribute("unique-identifier")?.Value;
      if (!string.IsNullOrWhiteSpace(preresult))
      {
#if DEBUG
        Log.Information($"EPUB uuid <<{preresult}>>");
#endif
        result = preresult?.GetUuidFromString();
      }
    }

    return result;
  }

  /// <summary>
  /// Что-то делаем с файлом
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <param name="f">Действие</param>
  /// <returns>Результат</returns>
  private static Guid DefAct(string file, Func<Stream, Guid>? f = null)
  {
    f = f ?? Default;
    Stream? stream = default;
    Guid? result = default;
#if DEBUG
    Log.Information($"Work with {file}");
#endif
    try
    {
      stream = File.OpenRead(file);
      result = f.Invoke(stream);
    }
    catch (Exception e)
    {
#if DEBUG
      Log.Information(e.Message);
#endif
      stream?.Close();
    }
    return (Guid)result;
  }

  static ExtractUuid()
  {
#if DEBUG
    string log_fn = $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}/genuuid.log";
    Log.Logger = new LoggerConfiguration().WriteTo.File(log_fn).CreateLogger();
#endif
    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
  }

  /// <summary>
  /// Извлечение (или создание) детерминированного UUID. Можно и из нескольких файлов
  /// </summary>
  /// <param name="args">Имена файлов</param>
  /// <summary>
  public static void Main(string[] args)
  {
    if (args.Length == 0)
    {
      Console.WriteLine("Извлечение (или создание) детерминированного UUID.");
    }
    else
    {
#if DEBUG
      Stopwatch sw = new();
      Log.Information("BEGIN");
      sw.Start();
#endif
      foreach (var file in args)
        try
        {
          if (File.Exists(file))
            Console.WriteLine($"{GetUuidByFN(file)}  {file}");
        }
        catch (Exception e)
        {
#if DEBUG
          Log.Warning(e.Message);
#endif
        }
#if DEBUG
      Log.Information($"END (TIME: {sw.ElapsedMilliseconds} ms)");
#endif
    }
  }

  /// <summary>
  /// Извлечение (или создание) детерминированного UUID
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  public static Guid GetUuidByFN(string file)
  {
    var ext = Path.GetExtension(file).ToLowerInvariant();
    Guid uuid;
    if(_opFuncDic.TryGetValue(ext, out Func<Stream, Guid> f))
    {
      uuid = DefAct(file, f);
    }
    else
    {
      uuid = DefAct(file, Default);
    }
    return uuid;
  }

  /// <summary>
  /// Извлечение UUID из PDF
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Результат</returns>
  /// <summary>
  public static Guid Pdf(Stream stream)
  {
    Guid? result = null;
    using (var pdf = new PdfReader(stream))
    {
      /* Исхожу из того, что в файле есть метаинформация в XMP. Насколько
       * я знаю UUID может быть в двух местах, но это знание получено
       * получением имеющихся PDF, поэтому никаких дополнительных проверок
       * не делаю, пусть лезут исключения - буду расширять свои знания. */
      var xmpbin = pdf.Metadata;
      if (xmpbin != null)
      {
        var _ = new UTF8Encoding().GetString(xmpbin);
        var xmlstr = _[0] != '\uFEFF' ? _ : _.Substring(1);
        XNamespace xapMm = "http://ns.adobe.com/xap/1.0/mm/";
        var xdoc = XDocument.Parse(xmlstr);
        var docId = xdoc.Descendants(xapMm + "DocumentID");
        if (docId.Count() > 0)
          result = docId.First().Value.GetUuidFromString();
        if (result == null)
        {
          XNamespace rdf = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
          try
          {
            var preresult = xdoc?.Descendants(rdf + "Description")
                ?.First(x => x.Attribute(xapMm + "DocumentID") != null)
                ?.Attribute(xapMm + "DocumentID")
                ?.Value;
#if DEBUG
            if (!string.IsNullOrWhiteSpace(preresult))
              Log.Information($"PDF DocumentID <<{preresult}>>");
#endif
            result = preresult?.GetUuidFromString();
          }
          catch (Exception e)
          {
#if DEBUG
            Log.Warning(e.Message);
#endif
          }
        }
      }
      //Если нет, то берём ID из трейлера
      if (result == null && pdf.Trailer.GetAsArray(PdfName.ID) != null)
      {
        var pdfarr = pdf.Trailer.GetAsArray(PdfName.ID);
        var preresult = ChangeByteOrder(pdfarr[0].GetBytes());
#if DEBUG
        if (preresult != null && preresult.Count() != 0)
          Log.Information($"PDF Trailer <<{BitConverter.ToString(preresult)}>>");
#endif
        result = new Guid(ChangeByteOrder(pdfarr[0].GetBytes()));
      }
    }
    //Если уж и здесь нет, считаем MD5
    if (result == null)
      result = Default(stream);
    return (Guid)result;
  }

  /// <summary>
  /// Извлечение UUID из FB2
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Результат</returns>
  public static Guid Fb2(Stream stream)
  {
    //Пробуем взять из /FictionBook/description/document-info/id
    try
    {
      XNamespace fb2 = "http://www.gribuser.ru/xml/fictionbook/2.0";
      var xdoc = XDocument.Load(stream);
      var fb2Id = xdoc?.Element(fb2 + "FictionBook")
          ?.Element(fb2 + "description")
          ?.Element(fb2 + "document-info")
          ?.Element(fb2 + "id")
          ?.Value;
      if (!string.IsNullOrWhiteSpace(fb2Id) && fb2Id.Length > 31)
      {
#if DEBUG
        Log.Information($"FB2 id <<{fb2Id}>>");
#endif
        if (Guid.TryParse(fb2Id, out Guid try1)) return try1;

        string trimmedString = fb2Id.Trim(TRIM_CHAR);
#if DEBUG
        Log.Information($"FB2 id2 <<{trimmedString}>>");
#endif
        if (trimmedString.Length > 31 && Guid.TryParse(trimmedString, out Guid try2))
          return try2;

        string cleanString = trimmedString.Replace("-", null);
#if DEBUG
        Log.Information($"FB2 id3 <<{cleanString}>>");
#endif
        if (cleanString.Length == 32 && Guid.TryParse(cleanString, out Guid try3))
          return try3;

        int firstHyphen = trimmedString.IndexOf('-');
        string? subString = null;
        if (firstHyphen > -1)
          subString = fb2Id.Substring(firstHyphen + 1).Replace("-", null);
#if DEBUG
        Log.Information($"FB2 id4 <<{subString}>>");
#endif
        if (subString?.Length == 32 && Guid.TryParse(subString, out Guid try4))
        {
          return try4;
        }

      }
    }
    catch (Exception e)
    {
#if DEBUG
      Log.Warning(e.Message);
#endif
    }
    return Default(stream);
  }

  /// <summary>
  /// Извлечение UUID из Epub
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Результат</returns>
  public static Guid Epub(Stream stream)
  {
    string? content_opf = null;
    try
    {
      content_opf = XmlInZip(stream, "META-INF/container.xml", Container);
    }
    catch (Exception e)
    {
#if DEBUG
      Log.Warning(e.Message);
#endif
    }
    if (content_opf != null)
    {
      try
      {
        var eid = XmlInZip(stream, content_opf, Content);
        if (eid != null)
         return (Guid)eid;
        else throw new Exception();
      }
      catch (Exception e)
      {
#if DEBUG
        Log.Warning(e.Message);
#endif
      }
    }
    return Default(stream);
  }

  /// <summary>
  /// Извлечение UUID из Docx
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Результат</returns>
  public static Guid Docx(Stream stream)
  {
    Guid? result = default;
    try
    {
      result = XmlInZip(stream, "word/settings.xml", Docx15);
    }
    catch (Exception e)
    {
#if DEBUG
      Log.Warning(e.Message);
#endif
    }
    if (result == default)
    {
      try
      {
        result = XmlInZip(stream, "docProps/custom.xml", DocxNS);
      }
      catch (Exception e)
      {
#if DEBUG
        Log.Warning(e.Message);
#endif
      }
    }
    if(result == default)
      result = Default(stream);
    return (Guid)result;
  }

  /// <summary>
  /// Получение UUID из файла любого типа
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Результат</returns>
  public static Guid Default(Stream stream)
  {
    stream.Seek(0, 0);
    using var md5 = MD5.Create();
    var sum = md5.ComputeHash(stream);
    var result = new Guid(ChangeByteOrder(sum));
#if DEBUG
    Log.Information($"Default md5sum <<{result}>>");
#endif
    return result;
  }

  /// <summary>
  /// Извлечение UUID из PDF
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  public static Guid Pdf(string file) => DefAct(file, Pdf);

  /// <summary>
  /// Извлечение UUID из FB2
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  public static Guid Fb2(string file) => DefAct(file, Fb2);

  /// <summary>
  /// Извлечение UUID из Epub
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  public static Guid Epub(string file) => DefAct(file, Epub);

  /// <summary>
  /// Извлечение UUID из Docx
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  public static Guid Docx(string file) => DefAct(file, Docx);

  /// <summary>
  /// Получение UUID из файла любого типа
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  public static Guid Default(string file) => DefAct(file, Default);
}