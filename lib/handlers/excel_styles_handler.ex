defmodule ExtrText.ExcelStylesHandler do
  @behaviour Saxy.Handler

  def handle_event(:start_element, {"numFmt", attributes}, state) when state.name == "numFmts" do
    {:ok, %{state | num_formats: [attributes | state.num_formats]}}
  end

  def handle_event(:start_element, {"xf", attributes}, state) when state.name == "cellXfs" do
    {:ok, %{state | cell_style_xfs: [attributes | state.cell_style_xfs]}}
  end

  def handle_event(:start_element, {_, _}, state) when state.name == "cellXfs" do
    {:ok, state}
  end

  def handle_event(:start_element, {name, _attributes}, state) do
    {:ok, %{state | name: name}}
  end

  def handle_event(_, _, state) do
    {:ok, state}
  end
end
