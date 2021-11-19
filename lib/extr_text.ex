defmodule ExtrText do
  @moduledoc """
  ExtrText is an Elixir library for extracting text and meta information from `.docx`, `.xlsx`,
  `.pptx` files.
  """

  @doc """
  Extract title, subject, description and body as a joined string from the specified binary data.
  """
  @spec extract(binary()) :: {:ok, String.t()}
  def extract(data) do
    tmpdir = System.tmp_dir!()
    {:ok, _paths} = :zip.unzip(data, cwd: tmpdir)

    attributes =
      case File.read(Path.join(tmpdir, "docProps/core.xml")) do
        {:ok, xml} -> extract_attributes(xml)
        {:error, _} -> nil
      end

    document =
      case File.read(Path.join(tmpdir, "word/document.xml")) do
        {:ok, xml} -> extract_text(xml)
        {:error, _} -> nil
      end

    {:ok, attributes <> "\n" <> document}
  end

  defp extract_attributes(xml) do
    {:ok, %{texts: texts}} = Saxy.parse_string(xml, AttributeHandler, %{name: nil, texts: []})
    reverse_and_join(texts)
  end

  defp extract_text(xml) do
    {:ok, texts} = Saxy.parse_string(xml, WordDocumentHandler, [])
    reverse_and_join(texts)
  end

  defp reverse_and_join(texts) do
    texts
    |> Enum.reverse()
    |> Enum.join("\n")
  end
end
