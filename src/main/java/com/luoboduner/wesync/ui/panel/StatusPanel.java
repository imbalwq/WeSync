package com.luoboduner.wesync.ui.panel;

import com.luoboduner.wesync.App;
import com.luoboduner.wesync.tools.ConstantsTools;
import com.luoboduner.wesync.tools.ExcelRead;
import com.luoboduner.wesync.ui.UiConsts;
import com.luoboduner.wesync.ui.component.MyIconButton;
import com.luoboduner.wesync.logic.ExecuteThread;
import com.luoboduner.wesync.vivo.bean.FactorTemplate;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.luoboduner.wesync.tools.PropertyUtil;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.ScheduledExecutorService;

/**
 * 状态面板
 *
 * @author lwq
 */
public class StatusPanel extends JPanel {

    private static final long serialVersionUID = 1L;
    private static final Logger logger = LoggerFactory.getLogger(StatusPanel.class);

    public static MyIconButton buttonStartSchedule;
    public static MyIconButton buttonStop;
    public static MyIconButton buttonStartNow;

//    public static JProgressBar progressTotal;
    public static JProgressBar progressCurrent;

    public static JLabel labelStatus;
    public static JLabel labelStatusDetail;
    public static JLabel labelFrom;
    //trados 报告log 存放目录
    public static JTextField textFieldLogForder;
    //trados 报告log 汇总输出位置
    public static JTextField textFieldLogForderOutput;
    //区分线下 线上report
    public static ButtonGroup buttonGroup;
    public static int buttonSelectValue;

    public static JLabel labelTo;
//    public static JLabel labelLastTime;
    public static JLabel labelKeepTime;
//    public static JLabel labelNextTime;
//    public static JLabel labelSuccess;
//    public static JLabel labelFail;
    private static JLabel labelLog;

    private static ScheduledExecutorService service;

    public static boolean isRunning = false;
    //头部
    public static JLabel labelTitle;

    /**
     * JComboBox<FactorTemplate> comboBox
     */

    public static JComboBox<FactorTemplate> comboBox;
    public static JLabel ftEdit;
    public static JLabel ftReload;

    /**
     * 构造
     */
    public StatusPanel() {
        super(true);
        initialize();
        addComponent();
        setContent();
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

        this.add(getUpPanel(), BorderLayout.NORTH);
        this.add(getCenterPanel(), BorderLayout.CENTER);
        this.add(getDownPanel(), BorderLayout.SOUTH);
    }

    /**
     * 上部面板
     *
     * @return
     */
    private JPanel getUpPanel() {
        JPanel panelUp = new JPanel();
        panelUp.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelUp.setLayout(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 5));

        labelTitle = new JLabel("Trados Log 汇总");
        labelTitle.setFont(UiConsts.FONT_TITLE);
        labelTitle.setForeground(UiConsts.TOOL_BAR_BACK_COLOR);
        panelUp.add(labelTitle);

        return panelUp;
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
        panelCenter.setLayout(new GridLayout(5, 1));

        // 状态Grid
        JPanel panelGridStatus = new JPanel();
        panelGridStatus.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelGridStatus.setLayout(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));

        labelStatus = new JLabel("执行结果");
        labelStatusDetail = new JLabel("详情：暂无生成记录");
        labelStatus.setFont(new Font(PropertyUtil.getProperty("ds.ui.font.family"), 1, 15));
        labelStatusDetail.setFont(UiConsts.FONT_NORMAL);

        labelStatus.setPreferredSize(UiConsts.LABLE_SIZE);
        labelStatusDetail.setPreferredSize(UiConsts.LABLE_SIZE);

        panelGridStatus.add(labelStatus);
        panelGridStatus.add(labelStatusDetail);

        // 来源/目标 Grid
        JPanel panelGridFromTo = new JPanel();
        panelGridFromTo.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelGridFromTo.setLayout(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP_15, 0));

        labelFrom = new JLabel();
        labelTo = new JLabel();
        labelFrom.setFont(UiConsts.FONT_NORMAL);
        labelTo.setFont(UiConsts.FONT_NORMAL);
//        labelFrom.setPreferredSize(UiConsts.LABLE_SIZE);
//        labelTo.setPreferredSize(UiConsts.LABLE_SIZE);
        labelFrom.setPreferredSize(UiConsts.LABLE_SIZE_ITEM2);
        labelTo.setPreferredSize(UiConsts.LABLE_SIZE_ITEM2);

        textFieldLogForder= new JTextField(80);
        textFieldLogForder.setPreferredSize(UiConsts.LABLE_SIZE);

        textFieldLogForderOutput = new JTextField(80);
        textFieldLogForderOutput.setPreferredSize(UiConsts.LABLE_SIZE);

        //如果设置了 使用默认生成位置 读取配置
        if("true".equals(ConstantsTools.CONFIGER.getAutoBak())){
            textFieldLogForderOutput.setText(ConstantsTools.CONFIGER.getMysqlPath());
        }


        //TODO 文件夹选择器 start
        //选择 Log 所在文件夹
        JButton openButton = new JButton("选择");
        openButton.setPreferredSize(UiConsts.LABLE_SIZE_ITEM);
        openButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int returnValue = fileChooser.showOpenDialog(null);

                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    String filePath = selectedFile.getAbsolutePath();
                    String fileName = selectedFile.getName();
                    System.out.println("Selected File Path: " + filePath);
                    System.out.println("Selected File Name: " + fileName);

                    textFieldLogForder.setText(filePath);
//                    System.exit(0); // 退出程序
                } else {
                    System.out.println("No file selected.");
                }
            }
        });

        JButton openButtonOutput = new JButton("选择");
        openButtonOutput.setPreferredSize(UiConsts.LABLE_SIZE_ITEM);
        openButtonOutput.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int returnValue = fileChooser.showOpenDialog(null);

                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    String filePath = selectedFile.getAbsolutePath();
                    textFieldLogForderOutput.setText(filePath);
//                    System.exit(0); // 退出程序
                } else {
                    System.out.println("No file selected.");
                }
            }
        });
        //TODO 文件夹选择器 end

        panelGridFromTo.add(labelFrom);
        panelGridFromTo.add(textFieldLogForder);
        panelGridFromTo.add(openButton);

        panelGridFromTo.add(labelTo);
        panelGridFromTo.add(textFieldLogForderOutput);
        panelGridFromTo.add(openButtonOutput);

        //

        // 详情Grid
        JPanel panelGridDetail = new JPanel();
        panelGridDetail.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelGridDetail.setLayout(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));

//        labelKeepTime = new JLabel();
        labelLog = new JLabel(PropertyUtil.getProperty("ds.ui.status.logDetail"));
//        labelKeepTime.setFont(UiConsts.FONT_NORMAL);
        labelLog.setFont(UiConsts.FONT_NORMAL);
//        labelKeepTime.setPreferredSize(new Dimension(300, 30));
        labelLog.setPreferredSize(UiConsts.LABLE_SIZE);
        labelLog.setForeground(UiConsts.TOOL_BAR_BACK_COLOR);

//        panelGridDetail.add(labelKeepTime);
//        panelGridDetail.add(labelSuccess);
//        panelGridDetail.add(labelFail);

        //查看日志详情
        panelGridDetail.add(labelLog);

        // 进度Grid
        JPanel panelGridProgress = new JPanel();
        panelGridProgress.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelGridProgress.setLayout(new GridLayout(2, 1, UiConsts.MAIN_H_GAP, 0));
        JPanel panelCurrentProgress = new JPanel();
        panelCurrentProgress.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelCurrentProgress.setLayout(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 20));
        JPanel panelTotalProgress = new JPanel();
        panelTotalProgress.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelTotalProgress.setLayout(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 0));

        JLabel labelCurrent = new JLabel(PropertyUtil.getProperty("ds.ui.status.progress.current"));
//        JLabel labelTotal = new JLabel(PropertyUtil.getProperty("ds.ui.status.progress.total"));
        labelCurrent.setFont(UiConsts.FONT_NORMAL);
//        labelTotal.setFont(UiConsts.FONT_NORMAL);
        progressCurrent = new JProgressBar();
//        progressTotal = new JProgressBar();

        Dimension preferredSizeLabel = new Dimension(80, 30);
        labelCurrent.setPreferredSize(preferredSizeLabel);
//        labelTotal.setPreferredSize(preferredSizeLabel);
        Dimension preferredSizeProgressbar = new Dimension(640, 20);
        progressCurrent.setPreferredSize(preferredSizeProgressbar);
//        progressTotal.setPreferredSize(preferredSizeProgressbar);

        panelCurrentProgress.add(labelCurrent);
        panelCurrentProgress.add(progressCurrent);
//        panelTotalProgress.add(labelTotal);
//        panelTotalProgress.add(progressTotal);

        panelGridProgress.add(panelCurrentProgress);
        panelGridProgress.add(panelTotalProgress);

        panelCenter.add(panelGridStatus);
        panelCenter.add(panelGridFromTo);
        panelCenter.add(panelGridDetail);
        panelCenter.add(panelGridProgress);

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
        panelDown.setLayout(new BoxLayout(panelDown, BoxLayout.X_AXIS));
//        panelDown.setLayout(new GridLayout(1, 2));
        JPanel panelGrid1 = new JPanel();
        panelGrid1.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelGrid1.setLayout(new FlowLayout(FlowLayout.LEFT, UiConsts.MAIN_H_GAP, 15));
        JPanel panelGrid2 = new JPanel();
        panelGrid2.setBackground(UiConsts.MAIN_BACK_COLOR);
        panelGrid2.setLayout(new FlowLayout(FlowLayout.RIGHT, UiConsts.MAIN_H_GAP, 15));

//        buttonStartSchedule = new MyIconButton(UiConsts.ICON_START_SCHEDULE, UiConsts.ICON_START_SCHEDULE_ENABLE,
//                UiConsts.ICON_START_SCHEDULE_DISABLE, "");
//        buttonStop = new MyIconButton(UiConsts.ICON_STOP, UiConsts.ICON_STOP_ENABLE,
//                UiConsts.ICON_STOP_DISABLE, "");
//        buttonStop.setEnabled(false);
//        buttonStartNow = new MyIconButton(UiConsts.ICON_SYNC_NOW, UiConsts.ICON_SYNC_NOW_ENABLE,
//                UiConsts.ICON_SYNC_NOW_DISABLE, "");

        //增加线上 线下选择按钮 start
        //end
        buttonStartNow = new MyIconButton(UiConsts.ICON_MAIN_START, UiConsts.ICON_MAIN_START_ENABLE, UiConsts.ICON_MAIN_START_DISABLE,"点击开始汇总您的Log文件");

//        panelGrid1.add(buttonStartSchedule);
//        panelGrid1.add(buttonStop);

        JPanel radioButtonPanel = new JPanel();
        radioButtonPanel.setBackground(UiConsts.MAIN_BACK_COLOR);
        radioButtonPanel.setPreferredSize(new Dimension(200, 30));
//        radioButtonPanel.setLayout(new FlowLayout(FlowLayout.LEFT, 10, 10));
        buttonGroup=new ButtonGroup();
        StatusPanel.buttonSelectValue=1;
        addRadioButton("常规Trados报告",1,radioButtonPanel);
        addRadioButton("Vivo线上报告",0,radioButtonPanel);

        //折算比例 start
        comboBox = new JComboBox<FactorTemplate>();
        //加载下拉列表数据
        reLoadComboBox();


        comboBox.setPreferredSize(new Dimension(200, 30));
        comboBox.setMaximumSize(new Dimension(200, 30));

//        comboBox.addActionListener(new ActionListener() {
//            @Override
//            public void actionPerformed(ActionEvent e) {
////                JOptionPane.showMessageDialog(App.statusPanel, "选中：" + comboBox.getSelectedItem(), PropertyUtil.getProperty("ds.ui.tips"),
////                        JOptionPane.CLOSED_OPTION);
////                panelDown.revalidate();
//            }
//        });

        comboBox.addItemListener(e -> {
            if (e.getStateChange() == ItemEvent.SELECTED) {
                Object selectedItem = e.getItem();
                //System.out.println("选择折算方案：" + selectedItem);
                comboBox.setPopupVisible(false);
                //JOptionPane.showMessageDialog(App.statusPanel, "选择折算方案：" + comboBox.getSelectedItem(), PropertyUtil.getProperty("ds.ui.tips"),
                //        JOptionPane.CLOSED_OPTION);
            }
        });
        //折算比例 end

        panelGrid2.add(buttonStartNow);
        radioButtonPanel.setPreferredSize(new Dimension(250, 30));
        radioButtonPanel.setMaximumSize(new Dimension(250, 30));

        JLabel labelZheSuan=new JLabel("折算方案 ");
        panelDown.add(radioButtonPanel);
        //折算比例 start
        panelDown.add(labelZheSuan);
        panelDown.add(comboBox);
         ftEdit=new JLabel(" [编辑]");
         ftReload=new JLabel(" [刷新]");

        ftEdit.setForeground(UiConsts.TOOL_BAR_BACK_COLOR);
        ftReload.setForeground(UiConsts.TOOL_BAR_BACK_COLOR);

        panelDown.add(ftEdit);
        panelDown.add(ftReload);
        //折算比例 end
        panelDown.add(panelGrid2);
        return panelDown;
    }

    /**
     * 重新加载 下拉框
     */
    private void reLoadComboBox() {
        //刷新数据时 先清空
        comboBox.removeAllItems();
        String factorTemplate= UiConsts.CURRENT_DIR + File.separator + "config" + File.separator + "factorTemplate.xlsx";
        String custFactorTemplate= UiConsts.CURRENT_DIR + File.separator + "config" + File.separator + "custFactorTemplate.xlsx";

        String factorData[][]= null;
        String custFactorData [][]=null;
        List<FactorTemplate> factorTemplateList = new ArrayList<>();
        try {
            factorData =ExcelRead.readExcelData(factorTemplate);
            custFactorData=ExcelRead.readExcelData(custFactorTemplate);
        }catch (Exception e){
            logger.error("加载折算比例数据异常",e);
            e.printStackTrace();
            JOptionPane.showMessageDialog(App.statusPanel, "加载折算比例数据异常:"+e.getMessage());
        }


        for(int i = 2; i< factorData.length; i++){
            if(StringUtils.isNotEmpty(factorData[i][0])){
                FactorTemplate ft=new FactorTemplate();
                ft.setCustomerName(factorData[i][0]);
                ft.setContextMatch(Double.valueOf(factorData[i][1]));
                ft.setRepetitions(Double.valueOf(factorData[i][2]));
                ft.setMatch_100(Double.valueOf(factorData[i][3]));
                ft.setMatch_95_99(Double.valueOf(factorData[i][4]));
                ft.setMatch_85_94(Double.valueOf(factorData[i][5]));
                ft.setMatch_75_84(Double.valueOf(factorData[i][6]));
                ft.setMatch_50_74(Double.valueOf(factorData[i][7]));
                ft.setMatch_new(Double.valueOf(factorData[i][8]));
                factorTemplateList.add(ft);
            }
        }

        for(int i = 2; i< custFactorData.length; i++){
            if(StringUtils.isNotEmpty(custFactorData[i][0])){
                FactorTemplate ft=new FactorTemplate();
                ft.setCustomerName(custFactorData[i][0]);
                ft.setContextMatch(Double.valueOf(custFactorData[i][1]));
                ft.setRepetitions(Double.valueOf(custFactorData[i][2]));
                ft.setMatch_100(Double.valueOf(custFactorData[i][3]));
                ft.setMatch_95_99(Double.valueOf(custFactorData[i][4]));
                ft.setMatch_85_94(Double.valueOf(custFactorData[i][5]));
                ft.setMatch_75_84(Double.valueOf(custFactorData[i][6]));
                ft.setMatch_50_74(Double.valueOf(custFactorData[i][7]));
                ft.setMatch_new(Double.valueOf(custFactorData[i][8]));
                factorTemplateList.add(ft);
            }
        }
        for (FactorTemplate template : factorTemplateList) {
            comboBox.addItem(template);
        }
    }

    /**
     * 设置状态面板组件内容
     */
    public static void setContent() {

        String logforder="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo LogReport\\log";
        String ouputforder="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo LogReport\\output";
        labelFrom.setText("Trados报告文件夹:");
        labelTo.setText("汇总文件存放位置: ");

        //labelKeepTime.setText(PropertyUtil.getProperty("ds.ui.status.keepTime") + ConstantsTools.CONFIGER.getLastKeepTime() + PropertyUtil.getProperty("ds.ui.status.second"));
//        labelSuccess.setText(PropertyUtil.getProperty("ds.ui.status.successTimes") + ConstantsTools.CONFIGER.getSuccessTime());
//        labelFail.setText(PropertyUtil.getProperty("ds.ui.status.failTimes") + ConstantsTools.CONFIGER.getFailTime());
    }

    /**
     * 为各组件添加事件监听
     */
    private void addListener() {
        buttonStartNow.addActionListener(e -> {
            if (isRunning == false) {
                buttonStartNow.setEnabled(false);
                //TODO 检查输入框位置是否为空
                String logforder=StatusPanel.textFieldLogForder.getText().trim();
                String outputforder=StatusPanel.textFieldLogForderOutput.getText().trim();
                if(StringUtils.isEmpty(logforder)){
                    JOptionPane.showMessageDialog(App.statusPanel, "请选择报告文件夹！", PropertyUtil.getProperty("ds.ui.tips"),
                            JOptionPane.WARNING_MESSAGE);
                    buttonStartNow.setEnabled(true);
                    return;
                }

                if(StringUtils.isEmpty(outputforder)){
                    JOptionPane.showMessageDialog(App.statusPanel, "请选择汇总文件存放位置！", PropertyUtil.getProperty("ds.ui.tips"),
                            JOptionPane.WARNING_MESSAGE);
                    buttonStartNow.setEnabled(true);
                    return;
                }

                StatusPanel.setContent();
                StatusPanel.progressCurrent.setValue(0);
//                labelStatus.setText(PropertyUtil.getProperty("Trados 报告正在汇总......"));
                ExecuteThread syncThread = new ExecuteThread();
                syncThread.start();
            }
        });

        labelLog.addMouseListener(new MouseListener() {

            @Override
            public void mouseReleased(MouseEvent e) {
                // TODO Auto-generated method stub

            }

            @Override
            public void mousePressed(MouseEvent e) {
                // TODO Auto-generated method stub

            }

            @Override
            public void mouseExited(MouseEvent e) {
                // TODO Auto-generated method stub

            }

            @Override
            public void mouseEntered(MouseEvent e) {
                labelLog.setCursor(new Cursor(Cursor.HAND_CURSOR));
            }

            @Override
            public void mouseClicked(MouseEvent e) {
                try {
                    Desktop.getDesktop().open(new File(ConstantsTools.PATH_LOG));
                } catch (IOException e1) {
                    // TODO Auto-generated catch block
                    logger.error(e1.toString());
                }
            }
        });

        //折算方案编辑按钮
        ftEdit.addMouseListener(new MouseListener() {

            @Override
            public void mouseReleased(MouseEvent e) {
            }
            @Override
            public void mousePressed(MouseEvent e) {
            }
            @Override
            public void mouseExited(MouseEvent e) {
            }
            @Override
            public void mouseEntered(MouseEvent e) {
                ftEdit.setCursor(new Cursor(Cursor.HAND_CURSOR));
            }

            @Override
            public void mouseClicked(MouseEvent e) {
                try {
                    Desktop.getDesktop().open(new File(ConstantsTools.PATH_CustFactorTemplate));
                } catch (IOException e1) {
                    // TODO Auto-generated catch block
                    logger.error(e1.toString());
                }
            }
        });

        //折算方案 刷新按钮
        ftReload.addMouseListener(new MouseListener() {

            @Override
            public void mouseReleased(MouseEvent e) {
            }
            @Override
            public void mousePressed(MouseEvent e) {
            }
            @Override
            public void mouseExited(MouseEvent e) {
            }
            @Override
            public void mouseEntered(MouseEvent e) {
                ftReload.setCursor(new Cursor(Cursor.HAND_CURSOR));
            }

            @Override
            public void mouseClicked(MouseEvent e) {
                try {
                    //刷新下拉列表
                    reLoadComboBox();
                } catch (Exception e1) {
                    // TODO Auto-generated catch block
                    logger.error(e1.toString());
                }
            }
        });
    }

    /**
     * 获取指定时间对应的毫秒数
     *
     * @param time "HH:mm:ss"
     * @return
     */
    private static long getTimeMillis(String time) {
        try {
            DateFormat dateFormat = new SimpleDateFormat("yy-MM-dd HH:mm:ss");
            DateFormat dayFormat = new SimpleDateFormat("yy-MM-dd");
            Date curDate = dateFormat.parse(dayFormat.format(new Date()) + " " + time);
            return curDate.getTime();
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return 0;
    }


    public void addRadioButton(String name, final int size,JPanel panel) {
        boolean selected=size==1;
        JRadioButton radioButton = new JRadioButton(name, selected);
        radioButton.setBackground(Color.WHITE);
        buttonGroup.add(radioButton);
        panel.add(radioButton);

        ActionListener listener=new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
//                JOptionPane.showMessageDialog(App.statusPanel, "当前选择："+size, PropertyUtil.getProperty("ds.ui.tips"),
//                        JOptionPane.PLAIN_MESSAGE);
                if (size == 0) {
                    labelTitle.setText("Vivo Log 汇总");
                }else if(size ==1 ){
                    labelTitle.setText("Trados Log 汇总");
                }
                buttonSelectValue=size;
            }
        };
        radioButton.addActionListener(listener);
    }
}
