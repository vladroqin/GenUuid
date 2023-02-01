using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using iTextSharp.text.pdf;
#if DEBUG
using Serilog;
#endif

namespace GenUuid;

public static class ExtractUuid
{
#if DEBUG
  private static string _log_fn =
      $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}/genuuid.log";
#endif

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
  private static T XmlInZip<T>(string file, string fileInArchive, Func<StreamReader, T> f)
  {
    using var stream = File.OpenRead(file);
    using var arc = new ZipArchive(stream, ZipArchiveMode.Read);
    using var sr = new StreamReader(arc.GetEntry(fileInArchive).Open());
#if DEBUG
    Log.Information(file);
#endif
    return f.Invoke(sr);
  }

  /// <summary>
  /// Извлечение UUID из файлов MS Word 2013+
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Результат</returns>
  private static Guid Docx15(StreamReader stream)
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
    return new Guid(eid);
  }

  /// <summary>
  /// Извлечение нестандартного UUID из Docx
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>Результат</returns>
  private static Guid Docx(StreamReader stream)
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
    return new Guid(eid);
  }

  /// <summary>
  /// Находим путь к файлу с информацией. Если его нет бросаем исключение
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
  /// Находим UUID
  /// </summary>
  /// <param name="stream">Поток</param>
  /// <returns>UUID</returns>
  private static Guid Content(StreamReader stream)
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

    return (Guid)result;
  }

  static ExtractUuid()
  {
#if DEBUG
    Log.Logger = new LoggerConfiguration().WriteTo.File(_log_fn).CreateLogger();
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
      foreach (var file in args)
        try
        {
          if (File.Exists(file))
            Console.WriteLine($"{GetUuidByFN(file)}  {file}");
        }
        catch (UnauthorizedAccessException e)
        {
          Console.Error.WriteLine(e.Message);
        }
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
    switch (ext)
    {
      case ".pdf":
        uuid = Pdf(file);
        break;
      case ".epub":
        uuid = Epub(file);
        break;
      case ".fbd":
      case ".fb2":
        uuid = Fb2(file);
        break;
      //case ".docm":
      /* Еще можно добавить попытку извлечения для других форматов МС Офиса,
       * но они либо его в готовом виде не содержат, либо встречаются крайне редко
       * (у меня).
       * Также можно добавить в будущем поддержку XPS, но по нему я дополнительно
       * информации найти не смог. */
      case ".docx":
        uuid = Docx(file);
        break;
      default:
        uuid = Default(file);
        break;
    }
    return uuid;
  }

  /// <summary>
  /// Извлечение UUID из ПДФ
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  /// <summary>
  public static Guid Pdf(string file)
  {
    Guid? result = null;
    using var stream = File.OpenRead(file);
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
              Log.Information($"PDF {file} DocumentID <<{preresult}>>");
#endif
            result = preresult?.GetUuidFromString();
          }
          catch (Exception e)
          {
#if DEBUG
            Log.Warning($"{file}: {e.Message}");
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
          Log.Information($"PDF {file} Trailer <<{BitConverter.ToString(preresult)}>>");
#endif
        result = new Guid(ChangeByteOrder(pdfarr[0].GetBytes()));
      }
    }
    //Если уж и здесь нет, считаем MD5
    if (result == null)
      result = Default(file);
    return (Guid)result;
  }

  /// <summary>
  /// Извлечение UUID из FB2
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  public static Guid Fb2(string file)
  {
    Guid? result = null;
    using var stream = File.OpenRead(file);
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
      //Собственно, тут возможны варианты, но я их пока не учитываю
      if (!string.IsNullOrWhiteSpace(fb2Id))
      {
#if DEBUG
        Log.Information($"FB2 {file} id <<{fb2Id}>>");
#endif
        result = Guid.Parse(fb2Id);
      }
    }
    catch (Exception e)
    {
#if DEBUG
      Log.Warning($"{file}: {e.Message}");
#endif
      result = Default(file);
    }
    return (Guid)result;
  }

  /// <summary>
  /// Извлечение UUID из Epub
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>UUID</returns>
  public static Guid Epub(string file)
  {
    string? content_opf = null;
    try
    {
      content_opf = XmlInZip(file, "META-INF/container.xml", Container);
    }
    catch (Exception e)
        when (e is not FileNotFoundException && e is not UnauthorizedAccessException)
    {
#if DEBUG
      Log.Warning($"{file}: {e.Message}");
#endif
    }
    if (content_opf != null)
    {
      try
      {
        var eid = XmlInZip(file, content_opf, Content);
        return eid;
      }
      catch (Exception e)
      {
#if DEBUG
        Log.Warning($"{file}: {e.Message}");
#endif
      }
    }
    return Default(file);
  }

  /// <summary>
  /// Извлечение UUID из Docx
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  public static Guid Docx(string file)
  {
    Guid? result = null;
    try
    {
      result = XmlInZip(file, "word/settings.xml", Docx15);
    }
    catch (Exception e)
        when (e is not FileNotFoundException && e is not UnauthorizedAccessException)
    {
#if DEBUG
      Log.Warning($"{file}: {e.Message}");
#endif
    }
    if (result == null)
    {
      try
      {
        result = XmlInZip(file, "docProps/custom.xml", Docx);
      }
      catch (Exception e)
      {
#if DEBUG
        Log.Warning($"{file}: {e.Message}");
#endif
        result = Default(file);
      }
    }
    return (Guid)result;
  }

  /// <summary>
  /// Получение UUID из любого файла
  /// </summary>
  /// <param name="file">Имя файла</param>
  /// <returns>Результат</returns>
  public static Guid Default(string file)
  {
    using var stream = File.OpenRead(file);
    using var md5 = MD5.Create();
    var sum = md5.ComputeHash(stream);
#if DEBUG
    Log.Information($"Default {file} md5sum");
#endif
    return new Guid(ChangeByteOrder(sum));
  }
}