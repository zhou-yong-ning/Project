var 线程ID
//从这里开始执行
function 执行()
    //从这里开始你的代码
    var time1 = 20
    sleep(time1 * 20)
    //获取表格数据标识码
    复制()
    sleep(time1)
    var BSM = getclipboard()
    //获取表格数据实地现状
    keydown(39)
    sleep(time1)
    复制()
    sleep(time1)
    var SDXZ = getclipboard()
    SDXZ = 字符串截取(SDXZ, 0, 字符串长度(SDXZ) - 2)
    //获取表格数据核查备注
    keydown(39)
    sleep(time1)
    复制()
    var HCBZ = getclipboard()
    //任务栏图标三
    mmlc(122, 1057, 0, time1, time1)
    //点击返回列表
    mmlc(225, 168, time1 * 3, time1, time1)
    //弹出搜索界面
    sleep(time1 * 20)
    //点击搜索框
    mmlc(700, 172, 0, time1, time1)
    全选()
    sleep(time1)
    //设置剪切版
    setclipboard(BSM)
    sleep(time1)
    粘贴()
    sleep(time1)
    //点击查询
    sleep(time1)
    keypress(13)
    //查询结果点击
    mousemove(1864, 322)
    var cr1 = getcolor(1873, 304, 1)
    while(cr1 != "f5f7fa")
        cr1 = getcolor(1873, 304, 1)
        sleep(200)
    end
    sleep(time1)
    if(getcolor(241, 373, 1) != "ffffff")
        sleep(300)
    else
        mouseleftclick(1)
    end
    //等待弹出窗口
    cr1 = getcolor(1690, 277, 1)
    while(cr1 != "409eff")
        cr1 = getcolor(1690, 277, 1)
        sleep(200)
    end
    sleep(500)
    //向下翻滚
    mmlc(1716, 375, 100, 100, 100)
    cr1 = getcolor(1781, 979, 1)
    while(cr1 != "66cc99")
        cr1 = getcolor(1781, 979, 1)
        mmlc(1716, 375, 0, 0, 100)
        sleep(200)
        keypress(34)
    end
    //    实地现状
    mmlc(1190, 581, time1 * 5, time1, time1 * 5)
    mmlc(1190, 630, time1 * 5, time1, time1 * 5)
    //    种植属性
    if(SDXZ == "耕地")
        mmlc(1541, 666, time1 * 5, time1, time1 * 5)
        mmlc(1541, 777, time1 * 5, time1, time1 * 5)
        mmlc(1134, 835, time1 * 5, time1, time1 * 5)
        mmlc(1134, 889, time1 * 5, time1, time1 * 5)
    else
        mmlc(1541, 666, time1 * 5, time1, time1 * 5)
        mmlc(1541, 712, time1 * 5, time1, time1 * 5)
        mmlc(1134, 835, time1 * 5, time1, time1 * 5)
        mmlc(1134, 920, time1 * 5, time1, time1 * 5)
    end
    //    种植作物
    mmlc(1513, 750, time1 * 5, time1, time1 * 5)
    keysendstring(字符串截取(HCBZ, 0, 字符串长度(HCBZ) - 2))
    mmlc(1513, 805, 500, time1, time1 * 5)
    mmlc(1413, 805, time1 * 20, time1, time1 * 5)
    //向下翻滚
    mmlc(1413, 805, 100, 100, 100)
    cr1 = getcolor(1781, 979, 1)
    while(cr1 != "66cc99")
        cr1 = getcolor(1781, 979, 1)
        mmlc(1413, 805, 0, 0, 100)
        sleep(200)
        keypress(34)
    end
    //    插入上报代码
    mmlc(1678, 980, 100, 100, 500)
    cr1 = getcolor(1856, 237, 1)
    while(cr1 != "f5f7fa")
        cr1 = getcolor(1856, 237, 1)
        sleep(200)
    end
    //任务栏图标四
    mmlc(170, 1057, 0, time1, time1 * 20)
    //    下
    keypress(40)
    //    左
    keypress(37, 2)
end
//控制脚本循环次数
function xunhuan()
    while(true)
        执行()
    end
end
//启动_热键操作
function 启动_热键()
    线程ID = threadbegin("xunhuan", "")
end
//终止热键操作
function 终止_热键()
    threadclose(线程ID)
end
function 复制()
    keydown(17)
    keypress(67)
    keyup(17)   
end
function 全选()
    keydown(17)
    keypress(65)
    keyup(17)   
end
function 粘贴()
    keydown(17)
    keypress(86)
    keyup(17)   
end
//鼠标移动点击
function mmlc(x, y, t1, t2, t3)
    sleep(t1)
    mousemove(x, y)
    sleep(t2)
    mouseleftclick(1)
    sleep(t3)
end