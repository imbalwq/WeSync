package com.luoboduner.wesync.vivo.bean;

/**
 * @author liweiqing
 * @date 2025/3/24 11:52
 * @description  折算比例
 */
public class FactorTemplate {
    //客户名称
    private String customerName;
    //完美匹配和上下文匹配
    private Double contextMatch;
    private Double repetitions;
    private Double match_100;
    private Double match_95_99;
    private Double match_85_94;
    private Double match_75_84;
    private Double match_50_74;
    private Double match_new;

    public String getCustomerName() {
        return customerName;
    }

    public void setCustomerName(String customerName) {
        this.customerName = customerName;
    }

    public Double getContextMatch() {
        return contextMatch;
    }

    public void setContextMatch(Double contextMatch) {
        this.contextMatch = contextMatch;
    }

    public Double getRepetitions() {
        return repetitions;
    }

    public void setRepetitions(Double repetitions) {
        this.repetitions = repetitions;
    }

    public Double getMatch_100() {
        return match_100;
    }

    public void setMatch_100(Double match_100) {
        this.match_100 = match_100;
    }

    public Double getMatch_95_99() {
        return match_95_99;
    }

    public void setMatch_95_99(Double match_95_99) {
        this.match_95_99 = match_95_99;
    }

    public Double getMatch_85_94() {
        return match_85_94;
    }

    public void setMatch_85_94(Double match_85_94) {
        this.match_85_94 = match_85_94;
    }

    public Double getMatch_75_84() {
        return match_75_84;
    }

    public void setMatch_75_84(Double match_75_84) {
        this.match_75_84 = match_75_84;
    }

    public Double getMatch_50_74() {
        return match_50_74;
    }

    public void setMatch_50_74(Double match_50_74) {
        this.match_50_74 = match_50_74;
    }

    public Double getMatch_new() {
        return match_new;
    }

    public void setMatch_new(Double match_new) {
        this.match_new = match_new;
    }


    @Override
        public String toString() {
            return customerName; // 显示名称
        }
}
