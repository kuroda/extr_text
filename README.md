# ExtrText

*ExtrText* is an Elixir library for extracting text and meta information from `.docx`, `.xlsx` and `.pptx` files.

## Usage

```elixir
docx = File.read!("example.docx")
{:ok, text} = ExtrText.extract(docx)
```
