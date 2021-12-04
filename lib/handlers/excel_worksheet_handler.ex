defmodule ExtrText.ExcelWorksheetHandler do
  @behaviour Saxy.Handler

  def handle_event(:start_element, {"row", _attributes}, state) do
    {:ok, %{state | buffer: []}}
  end

  def handle_event(:start_element, {"c", attributes}, state) do
    type =
      case Enum.find(attributes, fn {k, _v} -> k == "t" end) do
        {"t", v} -> v
        _ -> "n"
      end

    style =
      case Enum.find(attributes, fn {k, _v} -> k == "s" end) do
        {"s", v} -> parse_int(v)
        _ -> nil
      end

    {:ok, %{state | type: type, style: style}}
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
      case state.type do
        "s" -> format_string(chars, state)
        "n" -> format_number(chars, state)
        _ -> ""
      end

    {:ok, %{state | buffer: [string | state.buffer]}}
  end

  def handle_event(_, _, state) do
    {:ok, state}
  end

  defp format_string(chars, state) do
    if idx = parse_int(chars) do
      Enum.at(state.strings, idx)
    else
      ""
    end
  end

  defp format_number(chars, state) do
    pairs =
      if state.style && state.style > 0 do
        Enum.at(state.cell_style_xfs, state.style, [])
      else
        []
      end

    format_id = parse_int(get_value(pairs, "numFmtId"))

    if format_id do
      do_format_number(chars, format_id, state)
    else
      chars
    end
  end

  defp get_value(pairs, key) do
    pair =
      Enum.find(pairs, fn
        {k, _v} -> k == key
        _ -> false
      end)

    case pair do
      {_k, v} -> v
      _ -> nil
    end
  end

  defp do_format_number(chars, format_id, _state)
       when format_id in 14..36 or format_id in 45..47 or format_id in 50..58 do
    format_date_and_time(chars)
  end

  defp do_format_number(chars, format_id, state) do
    num_format =
      Enum.find(state.num_formats, fn pairs ->
        Enum.any?(pairs, fn
          {"numFmtId", v} -> v == Integer.to_string(format_id)
          _ -> false
        end)
      end)

    if num_format do
      format_code = get_value(num_format, "formatCode")

      if is_date_or_time?(format_code) do
        format_date_and_time(chars)
      else
        chars
      end
    else
      chars
    end
  end

  @date_format_words ~w(yyyy yy ggge ge mmm mm m dd d h ss)

  defp is_date_or_time?(format_code) do
    format_string = String.replace(format_code, ~r/\\./u, " ", global: true)
    words = Regex.scan(~r/[a-z]+/, format_string)
    words = List.flatten(words)

    Enum.any?(words, fn word -> word in @date_format_words end)
  end

  defp parse_int(nil), do: nil

  defp parse_int(str) do
    case Integer.parse(str) do
      {n, ""} -> n
      _ -> nil
    end
  end

  defp parse_float(nil), do: nil

  defp parse_float(str) do
    case Float.parse(str) do
      {n, ""} -> n
      _ -> nil
    end
  end

  defp format_date_and_time(chars) do
    i = parse_int(chars)

    if i do
      date = Date.add(~D[1899-12-30], i)
      Date.to_string(date)
    else
      f = parse_float(chars)

      if f do
        x = trunc(f)
        y = Decimal.sub(Decimal.from_float(f), Decimal.new(x))
        sec = Decimal.to_integer(Decimal.mult(y, 24 * 60 * 60))

        time =
          ~T[00:00:00]
          |> Time.add(sec)
          |> Time.truncate(:second)

        if x == 0 do
          Time.to_string(time)
        else
          date = Date.add(~D[1899-12-30], x)
          {:ok, dt} = DateTime.new(date, time, "Etc/UTC")
          String.trim_trailing(DateTime.to_string(dt), "Z")
        end
      else
        chars
      end
    end
  end
end
