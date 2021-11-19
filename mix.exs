defmodule ExtrText.MixProject do
  use Mix.Project

  def project do
    [
      app: :extr_text,
      version: "0.1.0",
      elixir: "~> 1.12",
      start_permanent: Mix.env() == :prod,
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
      {:ex_doc, "~> 0.25.0", only: :dev, runtime: false},
      {:saxy, "~> 1.4"}
    ]
  end
end
