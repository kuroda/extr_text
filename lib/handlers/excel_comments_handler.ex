defmodule ExtrText.ExcelCommentsHandler do
  @behaviour Saxy.Handler

  def handle_event(:start_element, {"commentList", _attributes}, _comments) do
    {:ok, []}
  end

  def handle_event(:characters, chars, comments) do
    {:ok, [chars | comments]}
  end

  def handle_event(:cdata, cdata, comments) do
    {:ok, [cdata | comments]}
  end

  def handle_event(_, _, comments) do
    {:ok, comments}
  end
end
