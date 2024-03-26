package com.xinyu.vivotool.vivo.util;

/**
 * @author liweiqing
 * @date 2023/9/19 16:31
 * @description
 */
public final class ExcelColumns {

    private ExcelColumns() {}

    private static String[] sources = new String[]{
            "A","B","C","D","E","F","G","H",
            "I","J","K","L","M","N","O","P",
            "Q","R","S","T","U","V","W","X","Y","Z"
    };

    /**
     * (256 for *.xls, 16384 for *.xlsx)
     * @param columnNum 列的个数，至少要为1
     * @throws IllegalArgumentException 如果 columnNum 超出该范围 [1,16384]
     * @return 返回[1,columnNum]共columnNum个对应xls列字母的数组
     */
    public static String[] getColumnLabels(int columnNum) {
        if(columnNum<1||columnNum>16384)
            throw new IllegalArgumentException();
        String[] columns = new String[columnNum];
        if(columnNum<27){	//小于27列 不用组合
            System.arraycopy(sources, 0, columns, 0, columnNum);
            return columns;
        }
        System.arraycopy(sources, 0, columns, 0, 26);	//前26列不需要进行组合

        //因为基于数组是从0开始，每到新一轮letterIdx 会递增，所以第一轮 在递增前是-1
        int letterIdx = -1;
        int currentLen = 26;//第一轮组合(2个字母的组合)是分别与A-Z进行拼接 所以是26
        int remainder;
        int lastLen = 0;	//用于定位上文提到的i%currentLen实际在数组中的位置
        int totalLen = 26;	//totalLen=currentLen+lastLen
        int currentLoopIdx = 0; //用来记录当前组合所有情形的个数

        for(int i=26;i<columnNum;i++){ //第27列(对应数组的第26个位置)开始组合

            //currentLen是上个组合所有情形的个数，与它取余找到要与上个组合的哪种情形进行拼接
            remainder = currentLoopIdx%currentLen;

            if(remainder==0){
                letterIdx++; //完成一次上个组合的遍历，转到下个字母进行拼接
                int j = letterIdx%26;

                //A-Z 26个子母都与上个组合所有情形都进行过拼接了，需要进行下个组合的拼接
                if(j==0&&letterIdx!=0){
                    lastLen = totalLen; //下个组合的lastLen是上个组合的totalLen

                    /**
                     * 下个组合的currentLen是上个组合的所有组合情形的个数
                     * （等于上个组合的currentLen*26)，26也就是拼接在前面的A-Z的个数
                     */
                    currentLen = 26*currentLen;

                    totalLen = currentLen+lastLen; //为下一轮的开始做准备
                    currentLoopIdx = 0; //到下一轮了 因此需要重置
                }
            }
            /**
             * sources[letterIdx%26]是该轮要拼接在前面的字母
             * columns[remainder+lastLen]是上个组合被拼接的情形
             */
            columns[i] = sources[letterIdx%26]+columns[remainder+lastLen];
            currentLoopIdx++;
        }
        return columns;
    }
}