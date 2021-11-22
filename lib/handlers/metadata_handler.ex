defmodule ExtrText.MetadataHandler do
  @behaviour Saxy.Handler

  def handle_event(:start_element, {name, _attributes}, state) do
    {:ok, %{state | name: name}}
  end

  def handle_event(:end_element, _name, state) do
    {:ok, %{state | name: nil}}
  end

  def handle_event(:characters, chars, %{name: "dc:title", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :title, chars)}}
  end

  def handle_event(:characters, chars, %{name: "dc:subject", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :subject, chars)}}
  end

  def handle_event(:characters, chars, %{name: "dc:description", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :description, chars)}}
  end

  def handle_event(:characters, chars, %{name: "dc:language", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :language, chars)}}
  end

  def handle_event(:characters, chars, %{name: "dc:creator", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :creator, chars)}}
  end

  def handle_event(:characters, chars, %{name: "cp:keywords", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :keywords, chars)}}
  end

  def handle_event(:characters, chars, %{name: "cp:lastModifiedBy", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :last_modified_by, chars)}}
  end

  def handle_event(:characters, chars, %{name: "cp:revision", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :revision, parse_int(chars))}}
  end

  def handle_event(:characters, chars, %{name: "dcterms:created", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :created, parse_datetime_string(chars))}}
  end

  def handle_event(:characters, chars, %{name: "dcterms:modified", metadata: metadata} = state) do
    {:ok, %{state | metadata: Map.put(metadata, :modified, parse_datetime_string(chars))}}
  end

  def handle_event(_, _, state) do
    {:ok, state}
  end

  defp parse_int(chars) do
    case Integer.parse(chars) do
      {i, ""} -> i
      _ -> nil
    end
  end

  defp parse_datetime_string(chars) do
    case DateTime.from_iso8601(chars) do
      {:ok, dt, _offset} -> dt
      {:error, _} -> nil
    end
  end
end
