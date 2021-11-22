defmodule ExtrTextTest do
  use ExUnit.Case

  @files_dir __DIR__ <> "/files"

  describe "extract/1" do
    test "extract text from a .docx file" do
      docx = File.read!(Path.join(@files_dir, "greeting.docx"))
      {:ok, text} = ExtrText.extract(docx)
      assert text == "Elixir\nTest\nGreeting\nGreeting\nHello, world!"
    end

    test "extract text from a .xlsx file" do
      xlsx = File.read!(Path.join(@files_dir, "prefectures.xlsx"))
      {:ok, text} = ExtrText.extract(xlsx)
      assert text == "Excel\nElixir\nNumbers\nPrefectures\nTokyo\nOsaka\nNagoya"
    end

    test "extract text from a .pptx file" do
      pptx = File.read!(Path.join(@files_dir, "hello.pptx"))
      {:ok, text} = ExtrText.extract(pptx)
      assert text == "spreadsheet\nHELLO\nPresentation\nHello, world!\nPage one\nFoo\nBar\nBaz"
    end
  end

  describe "get_metadata/1" do
    # The test file was created by LibreOffice Writer
    test "extract doc properties from a .docx file (1)" do
      docx = File.read!(Path.join(@files_dir, "greeting.docx"))
      {:ok, metadata} = ExtrText.get_metadata(docx)
      assert %ExtrText.Metadata{} = metadata
      assert metadata.title == "Greeting"
      assert metadata.subject == "Test"
      assert metadata.description == ""
      assert metadata.language == "ja-JP"
      assert metadata.keywords == "Elixir"
      assert metadata.creator == ""
      assert metadata.last_modified_by == ""
      assert metadata.revision == 1
      assert %DateTime{} = metadata.created
      assert %DateTime{} = metadata.modified
    end

    # The test file was created by Microsoft Word
    test "extract doc properties from a .docx file (2)" do
      docx = File.read!(Path.join(@files_dir, "greeting_jp.docx"))
      {:ok, metadata} = ExtrText.get_metadata(docx)
      assert %ExtrText.Metadata{} = metadata
      assert metadata.title == "Greeting"
      assert metadata.subject == ""
      assert metadata.description == "For test"
      assert metadata.language == ""
      assert metadata.keywords == "Elixir"
      assert metadata.creator == "黒田努"
      assert metadata.last_modified_by == "黒田努"
      assert metadata.revision == 3
      assert %DateTime{} = metadata.created
      assert %DateTime{} = metadata.modified
    end
  end
end
