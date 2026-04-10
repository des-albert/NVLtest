package org.dba.nvltest

import javafx.application.Application
import javafx.fxml.FXMLLoader
import javafx.scene.Scene
import javafx.scene.image.Image
import javafx.stage.Stage

class MainApplication : Application() {
    override fun start(stage: Stage) {
        val fxmlLoader = FXMLLoader(MainApplication::class.java.getResource("main.fxml"))
        val scene = Scene(fxmlLoader.load(), 1000.0, 640.0)
        val icon = Image(javaClass.getResourceAsStream("/img/NVLTest.png"))
        stage.icons.add(icon)
        stage.title = "NVL Config Check"
        stage.scene = scene
        stage.show()
    }
}

fun main() {
    Application.launch(MainApplication::class.java)
}

