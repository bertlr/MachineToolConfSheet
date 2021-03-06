/*
 * Copyright (C) 2016 by Herbert Roider <herbert@roider.at>
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package org.roiderh.machinetoolconfsheet;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
//import java.math.BigInteger;
import java.nio.charset.Charset;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.MissingResourceException;
import java.util.Set;
import java.util.HashSet;
import java.util.TreeMap;
import java.util.prefs.Preferences;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.JOptionPane;
import javax.swing.text.JTextComponent;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.netbeans.editor.BaseDocument;
import org.openide.awt.ActionID;
import org.openide.awt.ActionReference;
import org.openide.awt.ActionRegistration;
import org.openide.util.NbPreferences;
import org.netbeans.modules.editor.NbEditorUtilities;
import org.openide.filesystems.FileObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

@ActionID(
        category = "Tools",
        id = "org.roiderh.machinetoolconfsheet.CreateMachineToolConfSheetAction"
)
@ActionRegistration(
        iconBase = "org/roiderh/machinetoolconfsheet/hi-sheet24.png",
        displayName = "#CTL_CreateMachineToolConfSheetAction"
)
@ActionReference(path = "Toolbars/File", position = 0)
public final class CreateMachineToolConfSheetAction implements ActionListener {

    @Override
    public void actionPerformed(ActionEvent e) {
        BaseDocument doc = null;
        JTextComponent ed = org.netbeans.api.editor.EditorRegistry.lastFocusedComponent();
        if (ed == null) {
            JOptionPane.showMessageDialog(null, "Error: no open editor"); //NOI18N
            return;
        }

        FileObject fo = NbEditorUtilities.getFileObject(ed.getDocument());
        String path = fo.getPath();

        InputStream is = new ByteArrayInputStream(ed.getText().getBytes());

        BufferedReader br;

        br = new BufferedReader(new InputStreamReader(is, Charset.forName("UTF-8"))); //NOI18N

        ArrayList<String> lines = new ArrayList<>();

        try {
            String line;
            while ((line = br.readLine()) != null) {
                lines.add(line);
                System.out.println(line);
            }
        } catch (IOException x) {
            JOptionPane.showMessageDialog(null, "Error: " + x.getLocalizedMessage()); //NOI18N
        }

        TreeMap<Integer, Tool> tools = new TreeMap<>();
        ArrayList<String> programs = new ArrayList<>();
        Set<String> wks_dirs = new HashSet<String>();
        int activ_tool = -1;
        // Read all Tools with comments:
        for (int i = lines.size() - 1; i >= 0; i--) {
            String line = lines.get(i).trim();
            Matcher tool_change_command = Pattern.compile("(T)([0-9])+").matcher(line); //NOI18N

            if (line.startsWith("(") || line.startsWith(";")) { //NOI18N
                if (activ_tool >= 0) {

                    Tool t = tools.get(activ_tool);
                    if (line.startsWith("(")) { //NOI18N
                        line = line.substring(1, line.length() - 1);
                    } else {
                        line = line.substring(1, line.length());
                    }

                    t.text.add(line);
                    tools.put(activ_tool, t);
                }
            } else if (line.trim().startsWith("%")) { //NOI18N
                activ_tool = -1;
            } else if (tool_change_command.find()) {
                String ts = line.substring(tool_change_command.start() + 1, tool_change_command.end());
                activ_tool = Integer.parseInt(ts);
                if (!tools.containsKey(activ_tool)) {
                    tools.put(activ_tool, new Tool());
                }
            } else if (line.contains("M30") || line.contains("M17") || line.contains("M2") || line.contains("M02") || line.contains("RET")) { //NOI18N
                activ_tool = -1;

            } else {
                activ_tool = -1;
            }
            //System.out.println("line=" + line);
        }
        boolean is_header = false;
        ArrayList<String> header = new ArrayList<>();
        // Read the Comments at the beginning of the file:
        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i).trim();
            if (line.startsWith(";$PATH=/_N_")) { //NOI18N
                wks_dirs.add(this.extractPath(line));
            }
            if (line.startsWith("%")) { //NOI18N
                is_header = true;
                //programs.add(line.replaceAll(" ", "")); //NOI18N
                programs.add(this.parse_progname(line));

                //header.add(line.replaceAll(" ", "")); //NOI18N
            } else if (line.startsWith("(") || line.startsWith(";")) { //NOI18N
                if (is_header) {
                    if (line.startsWith("(")) { //NOI18N
                        line = line.substring(1, line.length() - 1);
                    } else {
                        line = line.substring(1, line.length());
                    }
                    if (line.startsWith("$PATH=/_N_") || line.length() < 1) { //NOI18N

                    } else {
                        header.add(line);
                    }

                }
            } else {
                is_header = false;
            }

        }

        try {
            Date dNow = new Date();
            SimpleDateFormat ft = new SimpleDateFormat("dd.MM.yyyy"); //NOI18N
            Date dModified = fo.lastModified();
            System.out.println("Current Date: " + ft.format(dNow)); //NOI18N
            InputStream in = CreateMachineToolConfSheetAction.class.getResourceAsStream("/org/roiderh/machinetoolconfsheet/resources/base_document.docx"); //NOI18N

            XWPFDocument document = new XWPFDocument(OPCPackage.open(in));
            //Write the Document in file system
            File tempFile = File.createTempFile("NcToolSettings_", ".docx"); //NOI18N
            try (FileOutputStream out = new FileOutputStream(tempFile)) {
                //XWPFTable table = document.getTableArray(0);

                XWPFParagraph title = document.getParagraphs().get(0);
                XWPFRun run = title.createRun();
                run.setText(org.openide.util.NbBundle.getMessage(CreateMachineToolConfSheetAction.class, "MachineToolConfSheet")); //NOI18N
                title = document.getParagraphs().get(1);
                run = title.createRun();
                run.setText(org.openide.util.NbBundle.getMessage(CreateMachineToolConfSheetAction.class, "Tools")); //NOI18N

                ArrayList<ArrayList<String>> table_text = new ArrayList<>();
                for (int i = 0; i < header.size(); i++) {

                    ArrayList<String> line = new ArrayList<>();
                    String name; // first column
                    String desc; // second column

                    int splitpos = header.get(i).indexOf(":");//NOI18N
                    if (splitpos > 1 && splitpos < 25) {
                        name = header.get(i).substring(0, splitpos).trim();
                        desc = header.get(i).substring(splitpos + 1);
                    } else {
                        name = "";//NOI18N
                        desc = header.get(i);
                    }
                    line.add(name);
                    line.add(desc);

                    table_text.add(line);

                }

                //XWPFTable table = document.createTable(table_text.size()+3, 2);
                //
                XWPFTable table = document.getTableArray(0);

                String wks_dir = "";
                for (Iterator i = wks_dirs.iterator(); i.hasNext();) {
                    if (wks_dir.length() > 0) {
                        wks_dir = wks_dir + ", "; //NOI18N
                    }
                    wks_dir = wks_dir + (String) i.next();

                }
                //wks_dir = String.join(", ", wks_dirs.toArray());
                XWPFTableRow tableRowHeader;
                String prog = String.join(", ", programs); //NOI18N
                table.getRow(0).getCell(0).setText(org.openide.util.NbBundle.getMessage(CreateMachineToolConfSheetAction.class, "ProgNr")); //NOI18N
                table.getRow(0).getCell(1).setText(prog);//NOI18N

                table.getRow(1).getCell(0).setText(org.openide.util.NbBundle.getMessage(CreateMachineToolConfSheetAction.class, "Workpiece")); //NOI18N
                table.getRow(1).getCell(1).setText(wks_dir);//NOI18N

                table.getRow(2).getCell(0).setText(org.openide.util.NbBundle.getMessage(CreateMachineToolConfSheetAction.class, "Filename")); //NOI18N
                table.getRow(2).getCell(1).setText(path);

                table.getRow(3).getCell(0).setText(org.openide.util.NbBundle.getMessage(CreateMachineToolConfSheetAction.class, "Date")); //NOI18N
                table.getRow(3).getCell(1).setText(ft.format(dNow));

                table.getRow(4).getCell(0).setText(org.openide.util.NbBundle.getMessage(CreateMachineToolConfSheetAction.class, "Modified")); //NOI18N
                table.getRow(4).getCell(1).setText(ft.format(dModified));

                //tableRowHeader = table.createRow();
                tableRowHeader = null;
                XWPFRun run_table;
                String prev_name = "dummy_1234567890sadfsaf"; //NOI18N
                for (int i = 0; i < table_text.size(); i++) {
                    String name = table_text.get(i).get(0);
                    String desc = table_text.get(i).get(1);

                    if (name.length() > 0) {
                        tableRowHeader = table.createRow();
                        tableRowHeader.getCell(0).getCTTc().addNewTcPr();
                        tableRowHeader.getCell(0).getCTTc().getTcPr().addNewTcBorders();
                        tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewBottom().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewLeft().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(1).getCTTc().addNewTcPr();
                        tableRowHeader.getCell(1).getCTTc().getTcPr().addNewTcBorders();
                        tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewBottom().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewLeft().setVal(STBorder.SINGLE);

                        run_table = tableRowHeader.getCell(1).addParagraph().createRun();
                        tableRowHeader.getCell(0).setText(name);
                        //tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.DASHED);
                        run_table.setText(desc);
                    } else if (prev_name.length() > 0 && name.length() == 0) {
                        tableRowHeader = table.createRow();
                        tableRowHeader.getCell(0).getCTTc().addNewTcPr();
                        tableRowHeader.getCell(0).getCTTc().getTcPr().addNewTcBorders();
                        tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewBottom().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewLeft().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(1).getCTTc().addNewTcPr();
                        tableRowHeader.getCell(1).getCTTc().getTcPr().addNewTcBorders();
                        tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewBottom().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewLeft().setVal(STBorder.SINGLE);

                        run_table = tableRowHeader.getCell(1).addParagraph().createRun();
                        tableRowHeader.getCell(0).setText("");   //NOI18N  

                        //tableRowHeader.getCell(1).setText(desc); 
                        run_table.setText(desc);
                    } else if (prev_name.length() == 0 && name.length() == 0) {
                        if (tableRowHeader == null) {
                            tableRowHeader = table.createRow();
                            tableRowHeader.getCell(0).getCTTc().addNewTcPr();
                            tableRowHeader.getCell(0).getCTTc().getTcPr().addNewTcBorders();
                            tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewBottom().setVal(STBorder.SINGLE);
                            tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.SINGLE);
                            tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewLeft().setVal(STBorder.SINGLE);
                            tableRowHeader.getCell(1).getCTTc().addNewTcPr();
                            tableRowHeader.getCell(1).getCTTc().getTcPr().addNewTcBorders();
                            tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewBottom().setVal(STBorder.SINGLE);
                            tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.SINGLE);
                            tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewLeft().setVal(STBorder.SINGLE);

                        }
                        run_table = tableRowHeader.getCell(1).addParagraph().createRun();
                        //run_table.addBreak();
                        run_table.setText(desc);
                        //tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.DASHED);
                    }
                    prev_name = name;
                }

                ///////////////////////////////////////////////////
                //table = document.getTableArray(1);
                //boolean first_line = true;
                Set keys = tools.keySet();
                //table = document.createTable(keys.size(), 2);
                table = document.getTables().get(1);
                //table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(700));
                //table.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(9000));
                XWPFTableRow newRow;
                for (int i = 0; i < keys.size(); i++) {
                    if (i > 0) {
                        tableRowHeader = table.createRow();
                        tableRowHeader.getCell(0).getCTTc().addNewTcPr();
                        tableRowHeader.getCell(0).getCTTc().getTcPr().addNewTcBorders();
                        tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewBottom().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(0).getCTTc().getTcPr().getTcBorders().addNewLeft().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(1).getCTTc().addNewTcPr();
                        tableRowHeader.getCell(1).getCTTc().getTcPr().addNewTcBorders();
                        tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewBottom().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewRight().setVal(STBorder.SINGLE);
                        tableRowHeader.getCell(1).getCTTc().getTcPr().getTcBorders().addNewLeft().setVal(STBorder.SINGLE);

                    }
                }
                int current_row = 0;
                for (Iterator i = keys.iterator(); i.hasNext();) {
                    int toolnr = (Integer) i.next();
                    Tool t = tools.get(toolnr);
                    XWPFTableRow tableRowTwo;

                    //current_row = table.getRows().size()-1;
                    tableRowTwo = table.getRow(current_row);
                    tableRowTwo.getCell(0).setText("T" + String.valueOf(toolnr)); //NOI18N

                    if (t.text.isEmpty()) {
                        tableRowTwo.getCell(1).addParagraph();
                    }
                    // The lines are in the reverse order, therfore reordering:
                    for (int j = t.text.size() - 1; j >= 0; j--) {
                        XWPFRun run_tool = tableRowTwo.getCell(1).addParagraph().createRun();
                        run_tool.setText(t.text.get(j));
                        if (j > 0) {
                            //run_tool.addBreak();
                            //run_tool.addCarriageReturn();
                        }
                    }
                    current_row++;
                }

                document.write(out);
            }
            System.out.println("create_table.docx written successully"); //NOI18N

            Runtime rt = Runtime.getRuntime();
            String os = System.getProperty("os.name").toLowerCase();//NOI18N
            String[] command = new String[2];
            //command[0] = "soffice";
            Preferences pref = NbPreferences.forModule(WordProcessingProgramPanel.class);
            command[0] = pref.get("executeable", "").trim();//NOI18N
            command[1] = tempFile.getCanonicalPath();
            File f = new File(command[0]);
            if (!f.exists()) {
                JOptionPane.showMessageDialog(null, "Error: program not found: " + command[0]); //NOI18N
                return;
            }

            Process proc = rt.exec(command); //NOI18N
            //System.out.println("ready created: " + tempFile.getCanonicalPath()); //NOI18N

        } catch (IOException | MissingResourceException ex) {
            JOptionPane.showMessageDialog(null, "Error: " + ex.getLocalizedMessage()); //NOI18N
        } catch (InvalidFormatException ife) {
            JOptionPane.showMessageDialog(null, "Error: " + ife.getLocalizedMessage()); //NOI18N
        }

    }

    String parse_progname(String line) {

        String progname = line.trim();
        if (progname.startsWith("%") == false) { //NOI18N
            return ""; //NOI18N
        }
        progname = progname.replaceAll(" ", ""); //NOI18N
        progname = progname.substring(1);
        if (progname.startsWith("MPF")) { //NOI18N
            progname = progname.substring(3);
            progname = progname.concat(".mpf"); //NOI18N
        } else if (progname.startsWith("SPF")) { //NOI18N
            progname = progname.substring(3);
            progname = progname.concat(".spf"); //NOI18N
        } else if (progname.startsWith("_N_")) { //NOI18N
            progname = progname.substring(3);
            if (progname.endsWith("_MPF_")) { //NOI18N
                progname = progname.substring(0, progname.length() - 5);
                progname = progname.concat(".mpf"); //NOI18N
            } else if (progname.endsWith("_SPF_")) { //NOI18N
                progname = progname.substring(0, progname.length() - 5);
                progname = progname.concat(".spf"); //NOI18N
            }

        }
        return progname;
    }

    private String extractPath(String line) {
        if (line.trim().startsWith(";$PATH=") == false) { //NOI18N
            return "";//NOI18N
        }
        String path = line.trim().substring(11);
        path = path.replace("_DIR/_N_", ".DIR/");//NOI18N
        path = path.substring(0, path.length() - 4) + ".WPD";//NOI18N

        return path;

    }
}
