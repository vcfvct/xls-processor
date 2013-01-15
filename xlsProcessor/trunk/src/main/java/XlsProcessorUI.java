import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.IOException;

import javax.swing.*;

/**
 * Created with IntelliJ IDEA. User: LiHa Date: 1/15/13
 */
public class XlsProcessorUI
{
    private JPanel panel1;
    private JButton startProcessButton;
    private JFileChooser fileChooser;

    public XlsProcessorUI()
    {
        fileChooser.setControlButtonsAreShown(false);
        fileChooser.setName("XLS Processor");
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.setFileFilter(new Xlsfilter());
        fileChooser.addPropertyChangeListener(new PropertyChangeListener() {
            public void propertyChange(PropertyChangeEvent evt) {
                if (JFileChooser.SELECTED_FILE_CHANGED_PROPERTY.equals(evt.getPropertyName())) {
                    File file = (File) evt.getNewValue();

                    if (file != null && file.isFile() && file.getName().contains("xls")) {
                        startProcessButton.setEnabled(true);

                    }
                    else if ( file != null ) {
                        startProcessButton.setEnabled(false);
                    }
                }

                fileChooser.repaint();
            }
        });


        startProcessButton.addActionListener(new ActionListener()
        {
            @Override
            public void actionPerformed(ActionEvent e)
            {
                File file = fileChooser.getSelectedFile();
                if (file == null)
                {
                    return;
                }
                ReturnGenerator generator = new ReturnGenerator(file);
                try
                {
                    generator.generate();
                }
                catch (IOException e1)
                {
                    e1.printStackTrace();
                }
            }
        });


    }

    public static void main(String[] args)
    {
        JFrame frame = new JFrame("XlsProcessorUI");
        frame.setContentPane(new XlsProcessorUI().panel1);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);

    }

}
