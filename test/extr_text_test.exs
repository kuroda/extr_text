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
  end
end
