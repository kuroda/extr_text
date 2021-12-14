defmodule ExtrText.ExcelDrawingsHandler do
  @behaviour Saxy.Handler

  def handle_event(:start_element, {"a:t", _attributes}, state) do
    {:ok, %{state | name: "a:t"}}
  end

  def handle_event(:end_element, _, state) do
    {:ok, %{state | name: nil}}
  end

  def handle_event(:characters, chars, state) when state.name == "a:t" do
    {:ok, %{state | texts: [chars | state.texts]}}
  end

  def handle_event(:cdata, cdata, state) when state.name == "a:t" do
    {:ok, %{state | texts: [cdata | state.texts]}}
  end

  def handle_event(_, _, texts) do
    {:ok, texts}
  end
end
