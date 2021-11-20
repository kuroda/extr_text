# ExtrText

[![ExtrText version](https://img.shields.io/hexpm/v/extr_text.svg)](https://hex.pm/packages/extr_text)
[![Hex.pm](https://img.shields.io/hexpm/dt/extr_text.svg)](https://hex.pm/packages/extr_text)

*ExtrText* is an Elixir library for extracting text and meta information from `.docx`, `.xlsx` and `.pptx` files.

## Usage

```elixir
docx = File.read!("example.docx")
{:ok, text} = ExtrText.extract(docx)

xlsx = File.read!("example.xlsx")
{:ok, text} = ExtrText.extract(xlsx)

pptx = File.read!("example.pptx")
{:ok, text} = ExtrText.extract(pptx)
```

## Installation

Add `:extr_text` to your `mix.exs`:

```elixir
  defp deps do
    [
      {:extr_text, "~> 0.1.0"}
    ]
end
```

Then, run `mix deps.get`.

## Acknowledgments

This project is inspired by [ranguba/chupa-text](https://github.com/ranguba/chupa-text),
a Ruby gem package.

## Author

[Tsutomu Kuroda](<mailto:t-kuroda@coregenik.com>)

## License

[MIT licens](./MIT_LICENSE.txt)
