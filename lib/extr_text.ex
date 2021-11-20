defmodule ExtrText do
  @moduledoc """
  ExtrText is an Elixir library for extracting text and meta information from `.docx`, `.xlsx`,
  `.pptx` files.
  """

  @doc """
  Extract title, subject, description and body as a joined string from the specified binary data.
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

  defp do_extract(data, subdir) do
    {:ok, paths} = :zip.unzip(data, cwd: subdir)

    type =
      cond do
        Enum.any?(paths, fn path -> path == subdir <> "/word/document.xml" end) -> :docx
        Enum.any?(paths, fn path -> path == subdir <> "/xl/sharedStrings.xml" end) -> :xsls
        true -> :unknown
      end

    attributes =
      case File.read(Path.join(subdir, "docProps/core.xml")) do
        {:ok, xml} -> extract_attributes(xml)
        {:error, _} -> nil
      end

    {handler, filename} =
      case type do
        :docs -> {WordDocumentHandler, "word/document.xml"}
        :xsls -> {ExcelSharedStringsHandler, "xl/sharedStrings.xml"}
        :unknown -> {nil, nil}
      end

    if handler && filename do
      document =
        case File.read(Path.join(subdir, filename)) do
          {:ok, xml} -> extract_text(handler, xml)
          {:error, _} -> nil
        end

      if attributes && document do
        {:ok, attributes <> "\n" <> document}
      else
        {:error, "Could not parse XML files."}
      end
    else
      {:error, "Could not find a target XML file."}
    end
  end

  defp extract_attributes(xml) do
    {:ok, %{texts: texts}} = Saxy.parse_string(xml, AttributeHandler, %{name: nil, texts: []})
    reverse_and_join(texts)
  end

  defp extract_text(handler, xml) do
    {:ok, texts} = Saxy.parse_string(xml, handler, [])
    reverse_and_join(texts)
  end

  defp reverse_and_join(texts) do
    texts
    |> Enum.reverse()
    |> Enum.join("\n")
  end
end
