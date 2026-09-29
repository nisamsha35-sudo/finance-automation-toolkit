// Power Query (M): clean a raw bank CSV. Synthetic data only.
// Change the file path to your copy of sample-data/bank_statement_sample.csv
let
    Source   = Csv.Document(File.Contents("C:\Path\bank_statement_sample.csv"),
                 [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Trimmed  = Table.TransformColumns(Promoted, {{"Description", Text.Trim, type text}}),
    Typed    = Table.TransformColumnTypes(Trimmed,
                 {{"Date", type date}, {"Debit", type number}, {"Credit", type number}}, "en-GB"),
    NoNulls  = Table.ReplaceValue(Typed, null, 0, Replacer.ReplaceValue, {"Debit","Credit"}),
    Amount   = Table.AddColumn(NoNulls, "Amount", each [Credit] - [Debit], type number),
    Result   = Table.RemoveColumns(Amount, {"Debit","Credit"})
in
    Result
