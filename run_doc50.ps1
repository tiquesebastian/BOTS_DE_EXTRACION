Set-Location "c:/Users/ticdesarrollo09/Music/trabajo/bot/EXTRATOR"
$env:DOC_TARGET_FILE = "c:/Users/ticdesarrollo09/Music/trabajo/bot/numeros_pendientes_50"
$env:DOC_OUTPUT_CSV = "c:/Users/ticdesarrollo09/Music/trabajo/bot/890701715_documentos_50.csv"
$env:DOC_OUTPUT_DETAIL = "c:/Users/ticdesarrollo09/Music/trabajo/bot/890701715_documentos_50_detalle.csv"
node extractor_documentos.js
