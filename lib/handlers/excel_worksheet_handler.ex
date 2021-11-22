defmodule ExtrText.ExcelWorksheetHandler do
  @behaviour Saxy.Handler

  def handle_event(:start_element, {"row", _attributes}, state) do
    {:ok, %{state | buffer: []}}
  end

  def handle_event(:start_element, {"c", attributes}, state) do
    type =
      case Enum.find(attributes, fn {k, _v} -> k == "t" end) do
        {"t", v} -> v
        _ -> nil
      end

    {:ok, %{state | type: type}}
  end

  def handle_event(:end_element, "row", state) do
    text =
      state.buffer
      |> Enum.reject(fn e -> e == "" end)
      |> Enum.reverse()
      |> Enum.join(" ")

    {:ok, %{state | texts: [text | state.texts]}}
  end

  def handle_event(:end_element, "c", state) do
    {:ok, %{state | type: nil}}
  end

  def handle_event(:characters, chars, state) do
    string =
      if state.type == "s" do
        if idx = parse_int(chars) do
          Enum.at(state.strings, idx)
        else
          ""
        end
      else
        ""
      end

    {:ok, %{state | buffer: [string | state.buffer]}}
  end

  def handle_event(_, _, state) do
    {:ok, state}
  end

  defp parse_int(str) do
    case Integer.parse(str) do
      {n, ""} -> n
      _ -> nil
    end
  end
end
