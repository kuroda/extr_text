defmodule ExtrText do
  @moduledoc """
  ExtrText is an Elixir library for extracting text and meta information from `.docx`, `.xlsx`,
  `.pptx` files.
  """

  @doc """
  Extracts properties (metadata) from the specified OOXML data.
  """
  @spec get_metadata(binary()) :: {:ok, ExtrText.Metadata.t()} | {:error, String.t()}
  def get_metadata(data) do
    case unzip(data) do
      {:ok, subdir, paths} -> do_get_metadata(subdir, paths)
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Extracts plain texts from the body of specified OOXML data.

  The return value is a double nested list of strings.

  Each element of outer list represents the sheets of `.xsls` data and the slides of `.pptx` data.
  For `.docx` data, the outer list has only one element.

  Each element of inner list represents the paragraphs or lines of a spreadsheet.
  """
  def get_texts(data) do
    case unzip(data) do
      {:ok, subdir, paths} -> do_get_texts(subdir, paths)
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Extracts comments from the body of specified OOXML data.

  The return value is a list of strings.

  Currently, only Excel files are supported.
  """
  @spec get_comments(binary()) :: {:ok, [String.t()]} | {:error, String.t()}
  def get_comments(data) do
    case unzip(data) do
      {:ok, subdir, paths} -> do_get_comments(subdir, paths)
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Extracts plain texts from draiwings embedded in the body of specified OOXML data.

  The return value is a list of strings.

  Currently, only Excel files are supported.
  """
  @spec get_texts_in_drawings(binary()) :: {:ok, [String.t()]} | {:error, String.t()}
  def get_texts_in_drawings(data) do
    case unzip(data) do
      {:ok, subdir, paths} -> do_get_texts_in_drawings(subdir, paths)
      {:error, reason} -> {:error, reason}
    end
  end

  defp unzip(data) do
    tmpdir = System.tmp_dir!()
    now = DateTime.utc_now()
    {usec, _} = now.microsecond
    subdir = tmpdir <> "/extr-text-" <> Integer.to_string(usec)

    case File.mkdir_p(subdir) do
      :ok -> do_unzip(data, subdir)
      {:error, _reason} -> {:error, "Can't create #{subdir}."}
    end
  end

  defp do_unzip(data, subdir) do
    case :zip.unzip(data, cwd: String.to_charlist(subdir)) do
      {:ok, paths} -> {:ok, subdir, Enum.map(paths, &List.to_string/1)}
      {:error, _reason} -> {:error, "Can't unzip the given data."}
    end
  end

  defp get_worksheets(subdir, paths) do
    Enum.filter(paths, fn path ->
      String.starts_with?(path, subdir <> "/xl/worksheets/") &&
        String.ends_with?(path, ".xml")
    end)
  end

  defp get_slides(subdir, paths) do
    Enum.filter(paths, fn path ->
      String.starts_with?(path, subdir <> "/ppt/slides/") &&
        String.ends_with?(path, ".xml")
    end)
  end

  defp do_get_metadata(subdir, _paths) do
    result =
      case File.read(Path.join(subdir, "docProps/core.xml")) do
        {:ok, xml} -> extract_metadata(xml)
        {:error, _} -> {:error, "Can't read docProps/core.xml."}
      end

    File.rm_rf!(subdir)
    result
  end

  defp extract_metadata(xml) do
    {:ok, %{metadata: metadata}} =
      Saxy.parse_string(xml, ExtrText.MetadataHandler, %{
        name: nil,
        metadata: %ExtrText.Metadata{}
      })

    {:ok, metadata}
  end

  defp do_get_texts(subdir, paths) do
    type = get_type(subdir, paths)
    result = do_get_texts(subdir, paths, type)
    File.rm_rf!(subdir)
    result
  end

  defp do_get_texts(_subdir, _paths, :unknown) do
    {:error, "Could not find a target Word/XML/PowerPoint file."}
  end

  defp do_get_texts(subdir, paths, :xlsx) do
    strings =
      if File.exists?(subdir <> "/xl/sharedStrings.xml") do
        ss_xml = File.read!(subdir <> "/xl/sharedStrings.xml")

        {:ok, strings} = Saxy.parse_string(ss_xml, ExtrText.ExcelSharedStringsHandler, [])
        Enum.reverse(strings)
      else
        []
      end

    st_xml = File.read!(subdir <> "/xl/styles.xml")

    {:ok, %{num_formats: num_formats, cell_style_xfs: cell_style_xfs}} =
      Saxy.parse_string(st_xml, ExtrText.ExcelStylesHandler, %{
        num_formats: [],
        cell_style_xfs: [],
        name: nil
      })

    num_formats = Enum.reverse(num_formats)
    cell_style_xfs = Enum.reverse(cell_style_xfs)

    worksheets = get_worksheets(subdir, paths)

    text_sets =
      worksheets
      |> Enum.map(fn path ->
        case File.read(path) do
          {:ok, xml} -> extract_texts(:xlsx, xml, strings, num_formats, cell_style_xfs)
          {:error, _} -> nil
        end
      end)
      |> Enum.reject(fn doc -> is_nil(doc) end)

    {:ok, text_sets}
  end

  defp do_get_texts(subdir, paths, type) when type in ~w(docx pptx)a do
    {handler, paths} =
      case type do
        :docx -> {ExtrText.WordDocumentHandler, [subdir <> "/word/document.xml"]}
        :pptx -> {ExtrText.PresentationSlideHandler, get_slides(subdir, paths)}
      end

    text_sets =
      paths
      |> Enum.map(fn path ->
        case File.read(path) do
          {:ok, xml} -> extract_texts(handler, xml)
          {:error, _} -> nil
        end
      end)
      |> Enum.reject(fn doc -> is_nil(doc) end)

    {:ok, text_sets}
  end

  defp extract_texts(:xlsx, xml, strings, num_formats, cell_style_xfs) do
    {:ok, %{texts: texts}} =
      Saxy.parse_string(xml, ExtrText.ExcelWorksheetHandler, %{
        texts: [],
        buffer: [],
        strings: strings,
        num_formats: num_formats,
        cell_style_xfs: cell_style_xfs,
        type: nil,
        style: nil
      })

    Enum.reverse(texts)
  end

  defp extract_texts(handler, xml) do
    {:ok, %{texts: texts}} = Saxy.parse_string(xml, handler, %{texts: [], buffer: []})
    Enum.reverse(texts)
  end

  defp do_get_comments(subdir, paths) do
    type = get_type(subdir, paths)
    result = extract_comments(subdir, type)
    File.rm_rf!(subdir)
    result
  end

  defp extract_comments(_subdir, :unknown) do
    {:error, "Could not find a target Word/XML/PowerPoint file."}
  end

  defp extract_comments(_subdir, type) when type in ~w(docx pptx)a do
    {:ok, []}
  end

  defp extract_comments(subdir, :xlsx) do
    comments =
      if File.exists?(subdir <> "/xl/comments1.xml") do
        c_xml = File.read!(subdir <> "/xl/comments1.xml")

        {:ok, strings} = Saxy.parse_string(c_xml, ExtrText.ExcelCommentsHandler, [])
        Enum.reverse(strings)
      else
        []
      end

    {:ok, comments}
  end

  defp do_get_texts_in_drawings(subdir, paths) do
    type = get_type(subdir, paths)
    result = extract_texts_in_drawings(subdir, type)
    File.rm_rf!(subdir)
    result
  end

  defp extract_texts_in_drawings(_subdir, :unknown) do
    {:error, "Could not find a target Word/XML/PowerPoint file."}
  end

  defp extract_texts_in_drawings(_subdir, type) when type in ~w(docx pptx)a do
    {:ok, []}
  end

  defp extract_texts_in_drawings(subdir, :xlsx) do
    texts =
      if File.exists?(subdir <> "/xl/drawings/drawing1.xml") do
        c_xml = File.read!(subdir <> "/xl/drawings/drawing1.xml")

        {:ok, %{texts: texts}} =
          Saxy.parse_string(c_xml, ExtrText.ExcelDrawingsHandler, %{name: nil, texts: []})

        Enum.reverse(texts)
      else
        []
      end

    {:ok, texts}
  end

  defp get_type(subdir, paths) do
    cond do
      Enum.any?(paths, fn path -> path == subdir <> "/word/document.xml" end) -> :docx
      Enum.any?(paths, fn path -> path == subdir <> "/xl/workbook.xml" end) -> :xlsx
      Enum.any?(paths, fn path -> path == subdir <> "/ppt/presentation.xml" end) -> :pptx
      true -> :unknown
    end
  end
end
