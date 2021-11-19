defmodule AttributeHandler do
  @behaviour Saxy.Handler

  def handle_event(:start_element, {name, _attributes}, state) do
    {:ok, %{state | name: name}}
  end

  def handle_event(:end_element, _name, state) do
    {:ok, %{state | name: nil}}
  end

  @names ~w(
    cp:keywords
    dc:description
    dc:subject
    dc:title
  )

  def handle_event(:characters, chars, %{name: name, texts: texts} = state) when name in @names do
    {:ok, %{state | texts: [chars | texts]}}
  end

  def handle_event(:cdata, chars, %{name: name, texts: texts} = state) when name in @names do
    {:ok, %{state | texts: [chars | texts]}}
  end

  def handle_event(_, _, state) do
    {:ok, state}
  end
end
