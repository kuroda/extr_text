defmodule ExtrText.MixProject do
  use Mix.Project

  def project do
    [
      app: :extr_text,
      version: "0.3.0",
      elixir: "~> 1.12",
      start_permanent: Mix.env() == :prod,
      description: "An Elixir library for extracting text from docs/xlsx/pptx files.",
      package: [
        maintainers: ["Tsutomu Kuroda"],
        licenses: ["MIT"],
        links: %{"GitHub" => "https://github.com/kuroda/extr_text"}
      ],
      deps: deps()
    ]
  end

  def application do
    [
      extra_applications: [:logger]
    ]
  end

  defp deps do
    [
      {:dialyxir, "~> 1.1.0", only: :dev, runtime: false},
      {:ex_doc, "~> 0.25.0", only: :dev, runtime: false},
      {:saxy, "~> 1.4"},
      {:decimal, "~> 2.0"}
    ]
  end
end
