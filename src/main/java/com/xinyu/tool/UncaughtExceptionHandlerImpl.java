package com.xinyu.tool;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class UncaughtExceptionHandlerImpl implements Thread.UncaughtExceptionHandler {
    private static final Logger logger = LoggerFactory.getLogger(UncaughtExceptionHandlerImpl.class);

    @Override
    public void uncaughtException(Thread t, Throwable e) {
        logger.error("未捕获的异常", e);
    }
}
