defmodule ExtrTextTest do
  use ExUnit.Case

  @files_dir __DIR__ <> "/files"

  describe "extract/1" do
    test "extract text from a .docx file" do
      docx = File.read!(Path.join(@files_dir, "greeting.docx"))
      {:ok, text} = ExtrText.extract(docx)
      assert text == "Elixir\nTest\nGreeting\nGreeting\nHello, world!"
    end
  end
end
