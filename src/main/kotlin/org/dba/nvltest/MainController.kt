package org.dba.nvltest

import javafx.application.Platform
import javafx.beans.property.SimpleStringProperty
import javafx.collections.FXCollections
import javafx.fxml.FXML
import javafx.scene.control.Button
import javafx.scene.control.Label
import javafx.scene.control.RadioButton
import javafx.scene.control.TableColumn
import javafx.scene.control.TableView
import javafx.scene.control.TableCell
import javafx.scene.control.ToggleGroup
import javafx.scene.text.Font
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.slf4j.Logger
import org.slf4j.LoggerFactory
import java.io.File
import java.io.IOException
import java.util.prefs.Preferences
import javafx.stage.FileChooser
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.SupervisorJob
import kotlinx.coroutines.javafx.JavaFx
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext


private val prefs: Preferences = Preferences.userNodeForPackage(MainController::class.java)
private val lastTargetDirKey = "lastTargetDir"
private lateinit var targetFile: File

private val sourceFileConfig = FileParseConfig(
    sheetName = "BoM",
    firstRow = 3,
    itemCol = 5,
    qtyCol = 6,
    skuCol = 7,
    solIDCol = 9
)

private val targetFileConfig = FileParseConfig(
    sheetName = "ExpertBOM",
    firstRow = 6,
    itemCol = 0,
    qtyCol = 1,
    skuCol = 2,
    solIDCol = 5
)


private val NVL72_GB300 = "NVL72 GB300 Configurator v18.0.xlsx"
private val NVL72_GB200 = "NVL72 GB200 Configurator v8.0.xlsx"
private val NVL4_GB200 = "NVL4 GB200 Configurator v10.0.xlsx"
private val NVL72_VR200 = "NVL72 VR200 Configurator v6.0.xlsx"

private val configRootPath = "C:\\Users\\albertd\\OneDrive - Hewlett Packard Enterprise\\HPE\\NVL\\"
private var sourcePath = configRootPath + NVL72_GB300
private var sourceFile = File(sourcePath)

private val logger: Logger = LoggerFactory.getLogger("Excel Reader")
private val dataFormatter = DataFormatter()
private val controllerScope = CoroutineScope(Dispatchers.JavaFx + SupervisorJob())

class MainController {
    @FXML
    lateinit var sourceSolIDCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var targetSolIDCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var skuCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var sourceQtyCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var targetQtyCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var sourceItemCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var targetItemCol: TableColumn<DiffResult, String>

    @FXML
    lateinit var tableDiff: TableView<DiffResult>

    @FXML
    lateinit var radio72VR200: RadioButton

    @FXML
    lateinit var radio4GB200: RadioButton

    @FXML
    lateinit var radio72GB200: RadioButton

    @FXML
    lateinit var confGroup: ToggleGroup

    @FXML
    lateinit var radio72GB300: RadioButton

    @FXML
    lateinit var buttonQuit: Button

    @FXML
    lateinit var buttonTarget: Button

    @FXML
    lateinit var buttonCompare: Button

    @FXML
    lateinit var labelTargetFile: Label

    @FXML
    lateinit var labelSourceFile: Label

    @FXML
    lateinit var labelJavaFX: Label

    @FXML
    lateinit var labelJDK: Label

    @FXML
    lateinit var labelStatus: Label

    @FXML
    fun initialize() {
        labelJDK.text = "Java SDK version: ${Runtime.version()}"
        labelJavaFX.text = "JavaFX version: ${System.getProperty("javafx.runtime.version")}"
        logger.info("Application started.")

        tableDiff.columns.setAll(
            makeColumn("SKU", 200.0, { it.sku }),
            makeColumn("Source Qty", 100.0, { it.sourceQty }) { it.sourceQty != it.targetQty },
            makeColumn("Target Qty", 100.0, { it.targetQty }) { it.sourceQty != it.targetQty },
            makeColumn("Source Item", 100.0, { it.sourceItem }) { it.sourceItem != it.targetItem },
            makeColumn("Target Item", 100.0, { it.targetItem }) { it.sourceItem != it.targetItem },
            makeColumn("Source SolID", 100.0, { it.sourceSolID }) { it.sourceSolID != it.targetSolID },
            makeColumn("Target SolID", 100.0, { it.targetSolID }) { it.sourceSolID != it.targetSolID }
        )
        tableDiff.isVisible = false
        labelSourceFile.text = sourceFile.name
    }

    private fun <T> makeColumn(
        title: String,
        width: Double,
        extractor: (DiffResult) -> T?,
        highlight: (DiffResult) -> Boolean = { false }
    ): TableColumn<DiffResult, String> {

        val col = TableColumn<DiffResult, String>(title)
        col.setCellValueFactory { cell ->
            SimpleStringProperty(extractor(cell.value)?.toString() ?: "")
        }

        col.prefWidth = width
        col.setCellFactory {
            object : TableCell<DiffResult, String>() {
                override fun updateItem(item: String?, empty: Boolean) {
                    super.updateItem(item, empty)

                    styleClass.removeAll("diff-cell", "match-cell")

                    if (empty) {
                        text = ""
                        return
                    }
                    val row = tableView.items[index]
                    text = item

                    if (highlight(row)) {
                        styleClass.add("diff-cell")
                    } else {
                        styleClass.add("match-cell")
                    }
                }
            }
        }
        return col
    }


    private fun readValuesFromFile(file: File, config: FileParseConfig): List<ValueList> {
        val values = mutableListOf<ValueList>()
        try {
            WorkbookFactory.create(file, null, true).use { workbook ->
                val evaluator = workbook.creationHelper.createFormulaEvaluator()
                val sheet = workbook.getSheet(config.sheetName)

                for (i in config.firstRow..sheet.lastRowNum) {
                    val row = sheet.getRow(i) ?: continue

                    val itemValue = dataFormatter.formatCellValue(row.getCell(config.itemCol), evaluator).trim()
                    if (itemValue.isBlank() || itemValue.equals("Total", ignoreCase = true)) continue

                    val qtyValue = dataFormatter.formatCellValue(row.getCell(config.qtyCol), evaluator).trim()
                    val skuValue = dataFormatter.formatCellValue(row.getCell(config.skuCol), evaluator).trim()
                    val solIDValue = dataFormatter.formatCellValue(row.getCell(config.solIDCol), evaluator).trim()

                    values.add(ValueList(itemValue, qtyValue, skuValue, solIDValue))
                }
            }
        } catch (e: IOException) {
            logger.error("Error reading file ${file.name}", e)
            throw e
        }

        return values
    }

    fun handleCompare() {

        tableDiff.isVisible = false
        controllerScope.launch {
            try {
                withContext(Dispatchers.JavaFx) {
                    val configValues = readValuesFromFile(sourceFile, sourceFileConfig)
                    labelStatus.text = "Source File Read"
                    logger.info("File ${sourceFile} Read ")

                    val targetValues = readValuesFromFile(targetFile, targetFileConfig)

                    if (configValues.size != targetValues.size) {
                        labelStatus.text = "Different size"
                        labelStatus.font = Font.font(36.0)
                        labelStatus.textFill = javafx.scene.paint.Color.RED
                    } else {
                        labelStatus.text = "Same size"
                        labelStatus.font = Font.font(24.0)
                        labelStatus.textFill = javafx.scene.paint.Color.GREEN
                    }

                    logger.info("File ${targetFile} Read ")

                    val mapTarget = targetValues.filter { it.sku != null }.associateBy { it.sku }

                    val diffList = configValues.mapNotNull { source ->
                        val target = mapTarget[source.sku]

                        when {
                            target == null -> DiffResult(
                                source.sku,
                                source.item, null,
                                source.quantity, null,
                                source.solID, null
                            )

                            source != target -> DiffResult(
                                source.sku,
                                source.item, target.item,
                                source.quantity, target.quantity,
                                source.solID, target.solID
                            )
                            else -> null
                        }
                    }
                    logger.info("Diffs Created ")
                    if (diffList.size > 0 ) {
                        tableDiff.isVisible = true
                        tableDiff.items = FXCollections.observableArrayList(diffList)
                    } else {
                        labelStatus.text = "File Match"
                        labelStatus.font = Font.font(24.0)
                        labelStatus.textFill = javafx.scene.paint.Color.GREEN
                    }
                }

            } catch (e: Exception) {
                logger.error("A critical error occurred during file processing.", e)
                labelStatus.text = "ERROR: An unexpected error occurred. Check logs"
            }
        }

    }

    fun handleOpenTargetFile() {
        val initialDir = prefs.get(lastTargetDirKey, System.getProperty("user.home"))
        openFileChooser(initialDir)?.let { file ->
            targetFile = file
            labelTargetFile.text = file.name
            prefs.put(lastTargetDirKey, file.parent) // Save the new directory
        }
    }

    private fun openFileChooser(directoryPath: String): File? {
        val fileChooser = FileChooser().apply {
            title = "Open Target Excel File"
            initialDirectory = File(directoryPath)
            extensionFilters.add(FileChooser.ExtensionFilter("Excel Files (*.xlsx)", "*.xlsx"))
        }
        return fileChooser.showOpenDialog(labelJavaFX.scene.window)
    }

    fun sourceSelect() {
        var selectedSource = ""
        val selectedToggle = confGroup.selectedToggle
        if (selectedToggle != null) {
            val selectedRadio = selectedToggle as RadioButton
            selectedSource = selectedRadio.text
        }

        when (selectedSource) {
            "NVL72 GB300" -> {
                sourcePath = configRootPath + NVL72_GB300
            }

            "NVL72 GB200" -> {
                sourcePath = configRootPath + NVL72_GB200
            }

            "NVL4 GB200" -> {
                sourcePath = configRootPath + NVL4_GB200
            }

            "NVL72 VR200" -> {
                sourcePath = configRootPath + NVL72_VR200
            }
        }
        sourceFile = File(sourcePath)
        labelSourceFile.text = sourceFile.name
    }

    fun handleQuit() {
        Platform.exit()
    }
}