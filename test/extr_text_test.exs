defmodule ExtrTextTest do
  use ExUnit.Case

  @files_dir __DIR__ <> "/files"

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
      assert metadata.revision == 2
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

  describe "get_texts/1" do
    test "extract plain texts from the body of a .docx file (1)" do
      docx = File.read!(Path.join(@files_dir, "greeting.docx"))
      {:ok, [texts]} = ExtrText.get_texts(docx)
      assert length(texts) == 2
      [line1, line2] = texts
      assert line1 == "Greeting"
      assert line2 == "Hello, world!"
    end

    test "extract plain texts from the body of a .docx file (2)" do
      docx = File.read!(Path.join(@files_dir, "greeting_jp.docx"))
      {:ok, [texts]} = ExtrText.get_texts(docx)
      assert length(texts) == 4
      [line1, line2, line3, line4] = texts
      assert line1 == "挨拶"
      assert line2 == ""
      assert line3 == "こんにちは。私は試験太郎です。"
      assert line4 == ""
    end

    # Note that this function does not extract numbers, dates, etc.
    test "extract plain texts from the slides of a .xlsx file" do
      xlsx = File.read!(Path.join(@files_dir, "prefectures.xlsx"))
      {:ok, [sheet1, sheet2]} = ExtrText.get_texts(xlsx)
      assert sheet1 == ["Tokyo", "Osaka", "Aichi"]
      assert sheet2 == ["Tokyo", "Osaka", "Fukuoka"]
    end

    test "extract plain texts from the slides of a .pptx file" do
      pptx = File.read!(Path.join(@files_dir, "hello.pptx"))
      {:ok, [slide1, slide2]} = ExtrText.get_texts(pptx)
      assert slide1 == ["Hello, world!"]
      assert slide2 == ["Page one", "Foo", "Bar", "Baz"]
    end
  end
end
