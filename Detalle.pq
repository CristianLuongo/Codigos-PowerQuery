// Detalle
let
    Origen = Folder.Files(Excel.CurrentWorkbook(){[Name="Carpeta"]}[Content]{0}[Column1]),
    #"Filas filtradas" = Table.SelectRows(Origen, each ([Extension] = ".txt") and Text.Contains([Name], "DETALLE", Comparer.OrdinalIgnoreCase)),
    #"Extraer datos" = Table.TransformColumns( #"Filas filtradas",{{"Content", each
let
Tabla = Csv.Document(_,[Delimiter =";",Encoding=1252]),
Transf = Table.SplitColumn(Tabla,"Column1", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false),{"Tipo Comprobante","Column1"}),
Separar = Table.SplitColumn(Transf, 
"Column1", 
Splitter.SplitTextByPositions({0, 8, 12, 20, 28, 40, 42, 57, 73, 88, 105, 109, 111, 200}),
 {"Fecha", "Sucursal", "Nro Comp Desde", "Nro Comp Hasta", "Cantidad", "Unidad de Medida", "Precio Unitario", "Importe Bonificado", "Bonificacion", "Total", "Tasa IVA", "Codigo", "Detalle"}
)
in
Separar
}}),
    #"Otras columnas quitadas" = Table.SelectColumns(#"Extraer datos",{"Content"}),
    #"Se expandió Content" = Table.ExpandTableColumn(#"Otras columnas quitadas", "Content", {"Tipo Comprobante", "Fecha", "Sucursal", "Nro Comp Desde", "Nro Comp Hasta", "Cantidad", "Unidad de Medida", "Precio Unitario", "Importe Bonificado", "Bonificacion", "Total", "Tasa IVA", "Codigo", "Detalle"}, {"Tipo Comprobante", "Fecha", "Sucursal", "Nro Comp Desde", "Nro Comp Hasta", "Cantidad", "Unidad de Medida", "Precio Unitario", "Importe Bonificado", "Bonificacion", "Total", "Tasa IVA", "Codigo", "Detalle"}),
    #"Tipo cambiado" = Table.TransformColumnTypes(#"Se expandió Content",{{"Tipo Comprobante", Int64.Type}, {"Fecha", type date}, {"Sucursal", type text}, {"Nro Comp Desde", type text}, {"Nro Comp Hasta", type text}, {"Cantidad", Int64.Type}, {"Unidad de Medida", Int64.Type}, {"Precio Unitario", Int64.Type}, {"Importe Bonificado", Int64.Type}, {"Bonificacion", Int64.Type}, {"Total", Int64.Type}, {"Tasa IVA", Int64.Type}, {"Codigo", type text}, {"Detalle", type text}}),
    #"Formato numeros" = Table.TransformColumns(#"Tipo cambiado",{
{"Precio Unitario", each _/100, type number},
{"Cantidad", each _/100000, type number},
{"Importe Bonificado", each _/100, type number},
{"Bonificacion", each _/100, type number},
{"Total", each _/1000, type number},
{"Tasa IVA", each _/100, type number},
{"Detalle", each Text.Trim(Text.Clean(Text.Upper(_))), type text}
}),
    Importe = Table.AddColumn( #"Formato numeros", "Importe", each if List.Contains({6,7,8,9,10,35,40,61,64,82},[Tipo Comprobante]) then [Cantidad] * ([Precio Unitario]/(1+([Tasa IVA]/100))) else [Cantidad] * [Precio Unitario], type number),
    #"Selección fecha" = Table.SelectRows(Importe, each [Fecha] >= List.Min(Periodo[Periodo]))
in
    #"Selección fecha"
