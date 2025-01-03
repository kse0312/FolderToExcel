module com.ns.foldertoexcel {
    requires javafx.controls;
    requires javafx.fxml;
    requires java.desktop;
    requires poi;
    requires poi.ooxml;

    opens com.ns.foldertoexcel to javafx.fxml;
    exports com.ns.foldertoexcel;
}