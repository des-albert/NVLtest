module org.dba {
    requires javafx.controls;
    requires javafx.fxml;
    requires kotlin.stdlib;
    requires org.slf4j;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires kotlinx.coroutines.core;
    requires kotlinx.coroutines.javafx;
    requires java.prefs;


    opens org.dba.nvltest to javafx.fxml;
    exports org.dba.nvltest;
}