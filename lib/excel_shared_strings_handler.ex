defmodule ExcelSharedStringsHandler do
  @behaviour Saxy.Handler

  def handle_event(:characters, chars, state) do
    {:ok, [chars | state]}
  end

  def handle_event(:cdata, cdata, state) do
    {:ok, [cdata | state]}
  end

  def handle_event(_, _, state) do
    {:ok, state}
  end
end
