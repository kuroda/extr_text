defmodule ExtrText.PresentationSlideHandler do
  @behaviour Saxy.Handler

  def handle_event(:start_element, {"a:p", _attributes}, state) do
    {:ok, %{state | buffer: []}}
  end

  def handle_event(:end_element, "a:p", state) do
    text =
      state.buffer
      |> Enum.reverse()
      |> Enum.join()

    {:ok, %{state | texts: [text | state.texts]}}
  end

  def handle_event(:characters, chars, state) do
    {:ok, %{state | buffer: [chars | state.buffer]}}
  end

  def handle_event(:cdata, cdata, state) do
    {:ok, %{state | buffer: [cdata | state.buffer]}}
  end

  def handle_event(_, _, state) do
    {:ok, state}
  end
end
