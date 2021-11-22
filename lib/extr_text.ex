defmodule ExtrText do
  @moduledoc """
  ExtrText is an Elixir library for extracting text and meta information from `.docx`, `.xlsx`,
  `.pptx` files.
  """

  @doc """
  Extract title, subject, description and body as a joined string from the specified binary data.

  The given data must be formatted in Office Open XML (OOXML).
  """
  @spec extract(binary()) :: {:ok, String.t()} | {:error, String.t()}
  def extract(data) do
    tmpdir = System.tmp_dir!()
    now = DateTime.utc_now()
    {usec, _} = now.microsecond
    subdir = tmpdir <> "/extr-text-" <> Integer.to_string(usec)

    case File.mkdir_p(subdir) do
      :ok -> do_extract(data, subdir)
      {:error, _reason} -> {:error, "Can't create #{subdir}."}
    end
  end

  @doc """
  Extract properties (metadata) from the specified OOXML data.
  """
  @spec get_metadata(binary()) :: {:ok, ExtrText.Metadata.t()} | {:error, String.t()}
  def get_metadata(data) do
    tmpdir = System.tmp_dir!()
    now = DateTime.utc_now()
    {usec, _} = now.microsecond
    subdir = tmpdir <> "/extr-text-" <> Integer.to_string(usec)

    case File.mkdir_p(subdir) do
      :ok -> do_get_metadata(data, subdir)
      {:error, _reason} -> {:error, "Can't create #{subdir}."}
    end
  end

  defp do_extract(data, subdir) do
    paths =
      case :zip.unzip(data, cwd: String.to_charlist(subdir)) do
        {:ok, paths} -> paths
        {:error, _} -> []
      end

    paths = Enum.map(paths, &List.to_string/1)

    type =
      cond do
        Enum.any?(paths, fn path -> path == subdir <> "/word/document.xml" end) -> :docx
        Enum.any?(paths, fn path -> path == subdir <> "/xl/sharedStrings.xml" end) -> :xslx
        Enum.any?(paths, fn path -> path == subdir <> "/ppt/presentation.xml" end) -> :pptx
        true -> :unknown
      end

    attributes =
      case File.read(Path.join(subdir, "docProps/core.xml")) do
        {:ok, xml} -> extract_attributes(xml)
        {:error, _} -> nil
      end

    {handler, paths} =
      case type do
        :docx -> {ExtrText.WordDocumentHandler, [subdir <> "/word/document.xml"]}
        :xslx -> {ExtrText.ExcelSharedStringsHandler, [subdir <> "/xl/sharedStrings.xml"]}
        :pptx -> {ExtrText.PresentationSlideHandler, get_slides(subdir, paths)}
        :unknown -> {nil, []}
      end

    result =
      if handler do
        documents =
          paths
          |> Enum.map(fn path ->
            case File.read(path) do
              {:ok, xml} -> extract_text(handler, xml)
              {:error, _} -> nil
            end
          end)
          |> Enum.reject(fn doc -> is_nil(doc) end)

        if attributes && length(documents) > 0 do
          {:ok, Enum.join([attributes | documents], "\n")}
        else
          {:error, "Could not parse XML files."}
        end
      else
        {:error, "Could not find a target XML file."}
      end

    File.rm_rf!(subdir)

    result
  end

  defp extract_attributes(xml) do
    {:ok, %{texts: texts}} =
      Saxy.parse_string(xml, ExtrText.AttributeHandler, %{name: nil, texts: []})

    reverse_and_join(texts)
  end

  defp extract_text(handler, xml) do
    {:ok, texts} = Saxy.parse_string(xml, handler, [])
    reverse_and_join(texts)
  end

  defp reverse_and_join(texts) do
    texts
    |> Enum.reverse()
    |> Enum.map(&String.trim/1)
    |> Enum.reject(fn text -> text == "" end)
    |> Enum.join("\n")
  end

  defp get_slides(subdir, paths) do
    Enum.filter(paths, fn path ->
      String.starts_with?(path, subdir <> "/ppt/slides/") &&
        String.ends_with?(path, ".xml")
    end)
  end

  defp do_get_metadata(data, subdir) do
    case :zip.unzip(data, cwd: String.to_charlist(subdir)) do
      {:ok, paths} -> paths
      {:error, _} -> []
    end

    case File.read(Path.join(subdir, "docProps/core.xml")) do
      {:ok, xml} -> extract_metadata(xml)
      {:error, _} -> {:error, "Can't read docProps/core.xml."}
    end
  end

  defp extract_metadata(xml) do
    {:ok, %{metadata: metadata}} =
      Saxy.parse_string(xml, ExtrText.MetadataHandler, %{
        name: nil,
        metadata: %ExtrText.Metadata{}
      })

    {:ok, metadata}
  end
end
