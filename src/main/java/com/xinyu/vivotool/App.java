package com.xinyu.vivotool;

import com.xinyu.vivotool.ui.UiConsts;
import com.xinyu.vivotool.ui.panel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;

/**
 * 程序入口，主窗口Frame
 *
 * @author lwq
 */
public class App {
    private static final Logger logger = LoggerFactory.getLogger(App.class);

    private JFrame frame;

    public static JPanel mainPanelCenter;

    public static StatusPanel statusPanel;
    public static SettingPanel settingPanel;

    /**
     * 程序入口main
     */
    public static void main(String[] args) {
//        EventQueue.invokeLater(() -> {
            try {
                App window = new App();
//                window.initialize();
//                StatusPanel.buttonStartSchedule.doClick();
                window.frame.setVisible(true);

            } catch (Exception e) {
                e.printStackTrace();
            }
//        });
    }

    public static void test(){
        JButton jButton = new JButton("a");
        jButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

            }
        });
        JButton jButton2 = new JButton("a");
        jButton2.addActionListener( (e)-> {
            System.out.println("asasas");
        });


    }

    /**
     * 构造，创建APP
     */
    public App() {
        initialize();
//        StatusPanel.buttonStartSchedule.doClick();
    }

    /**
     * 初始化frame内容
     */
    private void initialize() {


        //JOptionPane.showMessageDialog(App.statusPanel, "已过期！","",JOptionPane.ERROR_MESSAGE);
       // if(true)return;
        logger.info("==================Vivo工具 InitStart");
        // 设置系统默认样式
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException
                | UnsupportedLookAndFeelException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        // 初始化主窗口
        frame = new JFrame();
        frame.setBounds(UiConsts.MAIN_WINDOW_X, UiConsts.MAIN_WINDOW_Y, UiConsts.MAIN_WINDOW_WIDTH,
                UiConsts.MAIN_WINDOW_HEIGHT);
        frame.setTitle(UiConsts.APP_NAME);
        frame.setIconImage(UiConsts.IMAGE_ICON);
        frame.setBackground(UiConsts.MAIN_BACK_COLOR);
        JPanel mainPanel = new JPanel(true);
        mainPanel.setBackground(Color.white);
        mainPanel.setLayout(new BorderLayout());

        ToolBarPanel toolbar = new ToolBarPanel();
        statusPanel = new StatusPanel();



//        databasePanel = new DatabasePanel();
//        schedulePanel = new SchedulePanel();
//        backupPanel = new BackupPanel();
        settingPanel = new SettingPanel();

        mainPanel.add(toolbar, BorderLayout.WEST);
        mainPanelCenter = new JPanel(true);
        mainPanelCenter.setLayout(new BorderLayout());
        mainPanelCenter.add(statusPanel, BorderLayout.CENTER);

        mainPanel.add(mainPanelCenter, BorderLayout.CENTER);

        frame.add(mainPanel);

        frame.addWindowListener(new WindowListener() {

            @Override
            public void windowOpened(WindowEvent e) {
                // TODO Auto-generated method stub

            }

            @Override
            public void windowIconified(WindowEvent e) {
                // TODO Auto-generated method stub

            }

            @Override
            public void windowDeiconified(WindowEvent e) {
                // TODO Auto-generated method stub

            }

            @Override
            public void windowDeactivated(WindowEvent e) {
                // TODO Auto-generated method stub

            }

            @Override
            public void windowClosing(WindowEvent e) {
                if (StatusPanel.isRunning) {
                    JOptionPane.showMessageDialog(App.statusPanel,
                            "vivo相关文件正在处理，请等待执行完成！", "友情提示!", JOptionPane.WARNING_MESSAGE);
                } else {
                    frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
                }
            }

            @Override
            public void windowClosed(WindowEvent e) {
            }

            @Override
            public void windowActivated(WindowEvent e) {

            }
        });
        frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
        logger.info("==================Vivo工具 InitEnd");
    }
}
