# Changelog

## 0.3.2

* Extract functions as is from Excel sheets.
* Extract comments from Excel sheets.

## 0.3.1 (2021-12-04)

* Extract datetimes and times from Excel sheets.

## 0.3.0 (2021-12-04)

* Extract numbers and dates from Excel sheets.

## 0.2.1 (2021-11-22)

* Make private `ExtrText.do_unzip/1`.

## 0.2.0 (2021-11-22)

* Add `ExtrText.get_texts/1` for retrieving the content of OOXML document as a double nested list of strings.
* Add `ExtrText.get_metadata/1` for retrieving the properties of OOXML document as a struct.
* Remove `ExtrText.extract/1`.

## 0.1.0 (2021-11-20)
* Add `ExtrText.extract/1` for retrieving the content of OOXML document as a single text.
