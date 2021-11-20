# ExtrText

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

## Acknowledgments

This project is inspired by [ranguba/chupa-text](https://github.com/ranguba/chupa-text),
a Ruby gem package.

## Author

[Tsutomu Kuroda](<mailto:t-kuroda@coregenik.com>)

## License

[MIT licens](./MIT_LICENSE.txt)
