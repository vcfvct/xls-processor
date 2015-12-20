import java.awt.*;
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
    private JButton generateReturnButton;
    private JFileChooser fileChooser;
    private JProgressBar progressBar1;
    private JTextArea progressInfo;
    private JPanel buttonPanel;
    private JButton generateIRRButton;

    public XlsProcessorUI()
    {
        progressBar1.setStringPainted(true);
        progressBar1.setMinimum(0);
        progressBar1.setMaximum(100);
        progressInfo.setEnabled(false);
        fileChooser.setControlButtonsAreShown(false);
        fileChooser.setName("XLS Processor");
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.setFileFilter(new Xlsfilter());
        fileChooser.addPropertyChangeListener(evt -> {
            if (JFileChooser.SELECTED_FILE_CHANGED_PROPERTY.equals(evt.getPropertyName()))
            {
                File file = (File) evt.getNewValue();

                if (file != null && file.isFile() && file.getName().contains("xls"))
                {
                    generateReturnButton.setEnabled(true);
                    generateIRRButton.setEnabled(true);

                }
                else if (file != null)
                {
                    generateReturnButton.setEnabled(false);
                    generateIRRButton.setEnabled(false);
                }
            }

            fileChooser.repaint();
        });

        generateIRRButton.addActionListener(e -> {
            progressBar1.setValue(0);
            progressInfo.setText("");
            Runnable runner = () -> {
                File file = fileChooser.getSelectedFile();
                if (file == null)
                {
                    return;
                }
                IRRGenerator generator = new IRRGenerator(file);
                try
                {
                    generator.generate(progressBar1, progressInfo);
                }
                catch (IOException e1)
                {
                    e1.printStackTrace();
                }
            };
            Thread t = new Thread(runner, "Code Executer");
            t.start();
        });

        generateReturnButton.addActionListener(e -> {
            progressBar1.setValue(0);
            progressInfo.setText("");
            Runnable runner = () -> {
                File file = fileChooser.getSelectedFile();
                if (file == null)
                {
                    return;
                }
                ReturnGenerator generator = new ReturnGenerator(file);
                try
                {
                    generator.generate(progressBar1, progressInfo);
                }
                catch (IOException e1)
                {
                    e1.printStackTrace();
                }
            };
            Thread t = new Thread(runner, "Code Executer");
            t.start();
        });
    }

    public void updateBar(int newValue)
    {
        progressBar1.setValue(newValue);
    }

    public static void main(String[] args)
    {
        JFrame frame = new JFrame("XlsProcessor for KeelPoint -- Linda");
        frame.setPreferredSize(new Dimension(800,600));
        frame.setContentPane(new XlsProcessorUI().panel1);
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);

    }

}
