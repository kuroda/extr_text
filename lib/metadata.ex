defmodule ExtrText.Metadata do
  @moduledoc """
  This module defines a struct to hold properties (metadata) of an OOXML file.
  """

  defstruct title: "",
            subject: "",
            description: "",
            keywords: "",
            language: "",
            creator: "",
            last_modified_by: "",
            created: nil,
            modified: nil,
            revision: nil

  @type t :: %__MODULE__{
          title: String.t(),
          subject: String.t(),
          description: String.t(),
          keywords: String.t(),
          language: String.t(),
          creator: String.t(),
          last_modified_by: String.t(),
          created: DateTime.t() | nil,
          modified: DateTime.t() | nil,
          revision: non_neg_integer() | nil
        }
end
