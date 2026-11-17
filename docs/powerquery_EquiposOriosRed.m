// Consultas de Power Query del libro EquiposOriosRed.xlsm (extraídas del archivo original).
// Ambas leen el .txt que exporta Orios desde la carpeta local C:\7SAP\Export de la computadora donde se abre el libro.

section Section1;

shared UltimaToma = let
    Source = Csv.Document(File.Contents("C:\7SAP\Export\EquiposOrios.txt"),[Delimiter="	", Columns=12, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Planta", type text}, {"Area", type text}, {"Linea", type text}, {"Seccion", type text}, {"Fecha monitoreo", type datetime}, {"Equipo", type text}, {"Estatus", type text}, {"EstaActivo", type text}, {"Componente/Union", type text}, {"Sintoma", type text}, {"Recomendacion", type text}, {"Comentarios del Analista", type text}})
in
    #"Changed Type";

shared Todas = let
    Source = Csv.Document(File.Contents("C:\7SAP\Export\EquiposOrios.txt"),[Delimiter="	", Columns=12, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Planta", type text}, {"Area", type text}, {"Linea", type text}, {"Seccion", type text}, {"Fecha monitoreo", type datetime}, {"Equipo", type text}, {"Estatus", type text}, {"EstaActivo", type text}, {"Componente/Union", type text}, {"Sintoma", type text}, {"Recomendacion", type text}, {"Comentarios del Analista", type text}})
in
    #"Changed Type";
