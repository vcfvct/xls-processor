import java.io.File;

import javax.swing.filechooser.FileFilter;

/**
 * Created with IntelliJ IDEA. User: LiHa Date: 1/15/13
 */
class Xlsfilter extends FileFilter
{
    public boolean accept(File f) {
        return f.isDirectory() || f.getName().toLowerCase().endsWith(".xls") || f.getName().toLowerCase().endsWith(".xlsx");
    }

    public String getDescription() {
        return "Xls files";
    }
}
