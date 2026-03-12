package org.dba.nvltest

data class FileParseConfig(
    val sheetName: String,
    val firstRow: Int,
    val itemCol: Int,
    val qtyCol: Int,
    val skuCol: Int,
    val solIDCol: Int
)

data class ValueList(
    val item: String?,
    val quantity: String?,
    val sku: String?,
    val solID: String?
)

data class DiffResult(
    val sku: String?,
    val sourceItem: String?,
    val targetItem: String?,
    val sourceQty: String?,
    val targetQty: String?,
    val sourceSolID: String?,
    val targetSolID: String?,
)