package com.xinyu.vivotool.ui.panel;

import com.xinyu.vivotool.App;
import com.xinyu.vivotool.ui.UiConsts;
import com.xinyu.vivotool.ui.component.MyIconButton;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.xinyu.vivotool.tools.ConstantsTools;
import com.xinyu.vivotool.tools.PropertyUtil;

import javax.swing.*;
import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 高级选项面板
 *
 * @author lwq
 */
public class SettingPanelOption extends JPanel {

    private static final long serialVersionUID = 1L;

    private static MyIconButton buttonSave;

    private static MyIconButton buttionTableFiled;

    private static MyIconButton buttionClearLogs;


    /** vivo 译前 原文处理结果输出位置 */
    public static JCheckBox checkBox_VivoSplitDataInput;
    private static JTextField textField_VivoSplitDataInputForder;
    public static JCheckBox checkBox_VivoSplitDataOutput;
    private static JTextField textField_VivoSplitDataOutputForder;

    /** vivo 译后 原文处理结果输出位置 */
    private static JCheckBox checkBox_VivoMergeDataInput;
    private static JTextField textField_VivoMergeDataInputForder;
    private static JCheckBox checkBox_VivoMergeDataOutput;
    private static JTextField textField_VivoMergeDataOutputForder;

    private static final Logger logger = LoggerFactory.getLogger(SettingPanelOption.class);

    /**
     * 构造
     */
    public SettingPanelOption() {
        initialize();
        addComponent();
        setCurrentOption();
        addListener();
    }

    /**
     * 初始化
     */
    private void initialize() {
        this.setBackground(UiConsts.MAIN_BACK_COLOR);
        this.setLayout(new BorderLayout());
    }

    /**
     * 添加组件
     */
    private void addComponent() {

        this.add(getCenterPanel(), BorderLayout.CENTER);
        this.add(getDownPanel(), BorderLayout.SOUTH);

    }

    /**
     * 中部面板
     *
     * @return
     */
    private JPanel getCenterPanel() {
        // 中间面板
        JPanel panelCenter = new JPanel();
        panelCenter.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelCenter.setLayout(new GridLayout(1, 1));

        // 设置Grid
        JPanel panelGridOption = new JPanel();
        panelGridOption.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelGridOption.setLayout(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));

        // 初始化组件
        JPanel panelItem1 = new JPanel(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));
        JPanel panelItem2 = new JPanel(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));
        JPanel panelItem3 = new JPanel(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));
        JPanel panelItem4 = new JPanel(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));
        JPanel panelItem5 = new JPanel(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));
        JPanel panelItem6 = new JPanel(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));
        JPanel panelItem7 = new JPanel(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));
        JPanel panelItem8 = new JPanel(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));
        JPanel panelItem9 = new JPanel(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));

        panelItem1.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelItem2.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelItem3.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelItem4.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelItem5.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelItem6.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelItem7.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelItem8.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelItem9.setBackground(UiConsts.MAIN_BACK_COLOR);

        //清空Log
        panelItem1.setPreferredSize(UiConsts.PANEL_ITEM_SIZE);
        //译前 input、output
        panelItem2.setPreferredSize(UiConsts.PANEL_ITEM_SIZE);
        panelItem3.setPreferredSize(UiConsts.PANEL_ITEM_SIZE);
        panelItem4.setPreferredSize(new Dimension(1300, 40));
        panelItem5.setPreferredSize(UiConsts.PANEL_ITEM_SIZE);
        panelItem6.setPreferredSize(UiConsts.PANEL_ITEM_SIZE);
        panelItem7.setPreferredSize(UiConsts.PANEL_ITEM_SIZE);
        panelItem8.setPreferredSize(UiConsts.PANEL_ITEM_SIZE);
        panelItem9.setPreferredSize(UiConsts.PANEL_ITEM_SIZE);

//
        buttionClearLogs = new MyIconButton(UiConsts.ICON_CLEAR_LOG, UiConsts.ICON_CLEAR_LOG_ENABLE,
                UiConsts.ICON_CLEAR_LOG_DISABLE, "");
        panelItem1.add(buttionClearLogs);


        //TODO　译前　input、output
        checkBox_VivoSplitDataInput = new JCheckBox("[译前] 原文目录");
        checkBox_VivoSplitDataInput.setBackground(UiConsts.MAIN_BACK_COLOR);
        checkBox_VivoSplitDataInput.setFont(UiConsts.FONT_RADIO);
        panelItem2.add(checkBox_VivoSplitDataInput);
        textField_VivoSplitDataInputForder = new JTextField();
        textField_VivoSplitDataInputForder.setFont(UiConsts.FONT_RADIO);
        textField_VivoSplitDataInputForder.setPreferredSize(new Dimension(500, 26));
        panelItem3.add(textField_VivoSplitDataInputForder);

        checkBox_VivoSplitDataOutput = new JCheckBox("[译前] 拆分结果存储目录");
        checkBox_VivoSplitDataOutput.setBackground(UiConsts.MAIN_BACK_COLOR);
        checkBox_VivoSplitDataOutput.setFont(UiConsts.FONT_RADIO);
        panelItem4.add(checkBox_VivoSplitDataOutput);
        textField_VivoSplitDataOutputForder = new JTextField();
        textField_VivoSplitDataOutputForder.setFont(UiConsts.FONT_RADIO);
        textField_VivoSplitDataOutputForder.setPreferredSize(new Dimension(500, 26));
        panelItem5.add(textField_VivoSplitDataOutputForder);
//



        //TODO　译后　input、output
        checkBox_VivoMergeDataInput = new JCheckBox("[译后] 原文、及译文所在目录");
        checkBox_VivoMergeDataInput.setBackground(UiConsts.MAIN_BACK_COLOR);
        checkBox_VivoMergeDataInput.setFont(UiConsts.FONT_RADIO);
        panelItem6.add(checkBox_VivoMergeDataInput);
        textField_VivoMergeDataInputForder = new JTextField();
        textField_VivoMergeDataInputForder.setFont(UiConsts.FONT_RADIO);
        textField_VivoMergeDataInputForder.setPreferredSize(new Dimension(500, 26));
        panelItem7.add(textField_VivoMergeDataInputForder);

        checkBox_VivoMergeDataOutput = new JCheckBox("[译后] 合并后文件存储目录");
        checkBox_VivoMergeDataOutput.setBackground(UiConsts.MAIN_BACK_COLOR);
        checkBox_VivoMergeDataOutput.setFont(UiConsts.FONT_RADIO);
        panelItem8.add(checkBox_VivoMergeDataOutput);
        textField_VivoMergeDataOutputForder = new JTextField();
        textField_VivoMergeDataOutputForder.setFont(UiConsts.FONT_RADIO);
        textField_VivoMergeDataOutputForder.setPreferredSize(new Dimension(500, 26));
        panelItem9.add(textField_VivoMergeDataOutputForder);



        // 组合元素
        panelGridOption.add(panelItem1);
        panelGridOption.add(panelItem2);
        panelGridOption.add(panelItem3);
        panelGridOption.add(panelItem4);
        panelGridOption.add(panelItem5);
        panelGridOption.add(panelItem6);
        panelGridOption.add(panelItem7);
        panelGridOption.add(panelItem8);
        panelGridOption.add(panelItem9);

        panelCenter.add(panelGridOption);
        return panelCenter;
    }

    /**
     * 底部面板
     *
     * @return
     */
    private JPanel getDownPanel() {
        JPanel panelDown = new JPanel();
        panelDown.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelDown.setLayout(new FlowLayout(FlowLayout.RIGHT, UiConsts.MAIN_H_GAP, 15));

        buttonSave = new MyIconButton(UiConsts.ICON_SAVE, UiConsts.ICON_SAVE_ENABLE,
                UiConsts.ICON_SAVE_DISABLE, "");
        panelDown.add(buttonSave);

        return panelDown;
    }

    /**
     * 设置当前combox选项状态
     */
    public static void setCurrentOption() {

        //译前相关设置
        checkBox_VivoSplitDataInput.setSelected(Boolean.parseBoolean(ConstantsTools.CONFIGER.getXpathCheckboxSplitInputPath()));
        textField_VivoSplitDataInputForder.setText(ConstantsTools.CONFIGER.getXpathTextSplitInputPath());
        checkBox_VivoSplitDataOutput.setSelected(Boolean.parseBoolean(ConstantsTools.CONFIGER.getXpathCheckboxSplitOutputPath()));
        textField_VivoSplitDataOutputForder.setText(ConstantsTools.CONFIGER.getXpathTextSplitOutputPath());
        //译后相关设置
        checkBox_VivoMergeDataInput.setSelected(Boolean.parseBoolean(ConstantsTools.CONFIGER.getXpathCheckboxMergeInputPath()));
        textField_VivoMergeDataInputForder.setText(ConstantsTools.CONFIGER.getXpathTextMergeInputPath());
        checkBox_VivoMergeDataOutput.setSelected(Boolean.parseBoolean(ConstantsTools.CONFIGER.getXpathCheckboxMergeOutputPath()));
        textField_VivoMergeDataOutputForder.setText(ConstantsTools.CONFIGER.getXpathTextMergeOutputPath());

    }

    /**
     * 为相关组件添加事件监听
     */
    private void addListener() {
        buttonSave.addActionListener(e -> {

            try {
//                ConstantsTools.CONFIGER.setAutoBak(String.valueOf(checkBox_VivoSplitDataInput.isSelected()));
//                ConstantsTools.CONFIGER.setMysqlPath(textField_VivoSplitDataInputForder.getText());

                //译前 输入输出
                ConstantsTools.CONFIGER.setXpathCheckboxSplitInputPath(String.valueOf(checkBox_VivoSplitDataInput.isSelected()));
                ConstantsTools.CONFIGER.setXpathTextSplitInputPath(textField_VivoSplitDataInputForder.getText());
                ConstantsTools.CONFIGER.setXpathCheckboxSplitOutputPath(String.valueOf(checkBox_VivoSplitDataOutput.isSelected()));
                ConstantsTools.CONFIGER.setXpathTextSplitOutputPath(textField_VivoSplitDataOutputForder.getText());

                //译后 输入输出
                ConstantsTools.CONFIGER.setXpathCheckboxMergeInputPath(String.valueOf(checkBox_VivoMergeDataInput.isSelected()));
                ConstantsTools.CONFIGER.setXpathTextMergeInputPath(textField_VivoMergeDataInputForder.getText());
                ConstantsTools.CONFIGER.setXpathCheckboxMergeOutputPath(String.valueOf(checkBox_VivoMergeDataOutput.isSelected()));
                ConstantsTools.CONFIGER.setXpathTextMergeOutputPath(textField_VivoMergeDataOutputForder.getText());



                JOptionPane.showMessageDialog(App.settingPanel, PropertyUtil.getProperty("ds.ui.save.success"),
                        PropertyUtil.getProperty("ds.ui.tips"), JOptionPane.PLAIN_MESSAGE);
            } catch (Exception e1) {
                JOptionPane.showMessageDialog(App.settingPanel, PropertyUtil.getProperty("ds.ui.save.fail") + e1.getMessage(),
                        PropertyUtil.getProperty("ds.ui.tips"),
                        JOptionPane.ERROR_MESSAGE);
                logger.error("Write to xml file error" + e1.toString());
            }

        });

//        buttionTableFiled.addActionListener(e -> {
//            try {
//                Desktop.getDesktop().open(new File(ConstantsLogic.TABLE_FIELD_DIR));
//            } catch (IOException e1) {
//                logger.error("open table_field file fail:" + e1.toString());
//                e1.printStackTrace();
//            }
//
//        });

        buttionClearLogs.addActionListener(e -> {

            JOptionPane.showMessageDialog(App.settingPanel,"暂未启用log功能");
                    if(true)return;

            int answer = JOptionPane.showConfirmDialog(App.settingPanel,
                    PropertyUtil.getProperty("ds.ui.setting.clean.makeSure"),
                    PropertyUtil.getProperty("ds.ui.tips"), 2);

            if (answer == 0) {
                FileOutputStream testfile = null;
                try {
                    testfile = new FileOutputStream(ConstantsTools.PATH_LOG);
                    testfile.write(new String("").getBytes());
                    testfile.flush();
                    JOptionPane.showMessageDialog(App.settingPanel,
                            PropertyUtil.getProperty("ds.ui.setting.clean.success"),
                            PropertyUtil.getProperty("ds.ui.tips"),
                            JOptionPane.PLAIN_MESSAGE);
                } catch (IOException e1) {
                    JOptionPane.showMessageDialog(App.settingPanel,
                            PropertyUtil.getProperty("ds.ui.setting.clean.fail") + e1.getMessage(),
                            PropertyUtil.getProperty("ds.ui.tips"),
                            JOptionPane.ERROR_MESSAGE);
                    e1.printStackTrace();
                } finally {
                    if (testfile != null) {
                        try {
                            testfile.close();
                        } catch (IOException e1) {
                            JOptionPane.showMessageDialog(App.settingPanel,
                                    PropertyUtil.getProperty("ds.ui.setting.clean.fail") + e1.getMessage(),
                                    PropertyUtil.getProperty("ds.ui.tips"), JOptionPane.ERROR_MESSAGE);
                            e1.printStackTrace();
                        }
                    }
                }
            }
        });
    }
}
