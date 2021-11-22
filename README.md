# ExtrText

[![ExtrText version](https://img.shields.io/hexpm/v/extr_text.svg)](https://hex.pm/packages/extr_text)
[![Hex.pm](https://img.shields.io/hexpm/dt/extr_text.svg)](https://hex.pm/packages/extr_text)

*ExtrText* is an Elixir library for extracting text and meta information from `.docx`, `.xlsx` and `.pptx` files.

## Usage

```elixir
iex> docx = File.read!("example.docx")
iex> {:ok, texts} = ExtrText.get_texts(docx)
iex> texts
[
  ["Paragraph 1", "Paragraph 2", "Paragraph 3"]
]
iex> {:ok, metadata} = ExtrText.get_metadata(docx)
iex> metadata
%ExtrText.Metadata{
  created: ~U[2021-11-19 22:25:20Z],
  creator: "John Doe",
  description: "",
  keywords: "",
  language: "ja-JP",
  last_modified_by: "John Doe",
  modified: ~U[2021-11-22 21:24:43Z],
  revision: 2,
  subject: "",
  title: "Example"
}
```

## Installation

Add `:extr_text` to your `mix.exs`:

```elixir
  defp deps do
    [
      {:extr_text, "~> 0.2.0"}
    ]
end
```

Then, run `mix deps.get`.

## Limitations

* The function `ExtrText.get_texts/1` extracts texts from an Excel file, but not numbers, dates, etc.

## Acknowledgments

This project is inspired by [ranguba/chupa-text](https://github.com/ranguba/chupa-text),
a Ruby gem package.

## Author

[Tsutomu Kuroda](<mailto:t-kuroda@coregenik.com>)

## License

[MIT licens](./MIT_LICENSE.txt)
