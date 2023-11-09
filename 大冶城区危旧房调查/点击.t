var 线程ID
var time1 = 30
function 执行nc()
    var y = 396
    //从这里开始你的代码
    for(var i = 0; i < 10; i++)
        y = y + 39
        var cr0 = getcolor(1766, y, 1)
        if(cr0 == "068996")
        else
            y = y + 18
        end
        mmlc(1766, y, time1 * 3, time1 * 3, time1 * 3)
        var cr1 = getcolor(1478, 287, 1)
        while(cr1 != "eeeeee")
            cr1 = getcolor(1478, 287, 1)
            sleep(200)
        end 
        //        点击使用情况\安全信息
        mmlc(1296, 226, time1 * 10, time1 * 3, time1 * 3)
        cr1 = getcolor(1456, 330, 1)
        while(cr1 != "eeeeee")
            cr1 = getcolor(1456, 330, 1)
            sleep(200)
        end 
        mmlc(1587, 226, time1 * 3, time1 * 3, time1 * 3)
        cr1 = getcolor(1462, 349, 1)
        while(cr1 != "009caa")
            cr1 = getcolor(1280, 349, 1)
            sleep(200)
        end 
        //        返回列表
        mmlc(1876, 232, time1 * 20, time1 * 3, 0)
        cr1 = getcolor(1180, 277, 1)
        while(cr1 != "068996")
            cr1 = getcolor(1180, 277, 1)
            sleep(200)
        end 
        sleep(time1 * 2)
    end
end
//从这里开始执行
function 执行cz()
    var y = 396
    //从这里开始你的代码
    for(var i = 0; i < 10; i++)
        y = y + 39
        var cr0 = getcolor(1802, y, 1)
        if(cr0 == "068996")
        else
            y = y + 18
        end
        mmlc(1802, y, time1 * 3, time1 * 3, time1 * 3)
        var cr1 = getcolor(1478, 287, 1)
        while(cr1 != "eeeeee")
            cr1 = getcolor(1478, 287, 1)
            sleep(200)
        end 
        //        点击使用情况\安全信息
        mmlc(1296, 226, time1 * 3, time1 * 3, time1 * 3)
        cr1 = getcolor(1280, 290, 1)
        while(cr1 != "009caa")
            cr1 = getcolor(1280, 290, 1)
            sleep(200)
        end 
        mmlc(1403, 226, time1 * 3, time1 * 3, time1 * 3)
        cr1 = getcolor(1280, 349, 1)
        while(cr1 != "009caa")
            cr1 = getcolor(1280, 349, 1)
            sleep(200)
        end 
        //        返回列表
        mmlc(1876, 232, time1 * 20, time1 * 3, 0)
        cr1 = getcolor(1180, 277, 1)
        while(cr1 != "068996")
            cr1 = getcolor(1180, 277, 1)
            sleep(200)
        end 
        sleep(time1 * 2)
    end
end
function xunhuancz()
    var i0 = btn_点击()
    messagebox(i0)
    for(var i = i0; i <= i0 + 25; i++)
        执行cz()
        // 点击下一页
        // 区域找色
        var x = -1, y = -1
        var ret = findcolor(1784, 879, 1825, 1030, "57BEC7-000000", 0.8, 0, x, y)
        //        点击
        mmlc(x, y, time1, time1, time1 * 5)
        mousemove(1874, 425)
        var cr2 = getcolor(1874, 425, 1)
        while(cr2 != "ececec")
            cr2 = getcolor(1874, 425, 1)
            sleep(200)
        end
        //        下载数据包
        mmlc(164, 1055, time1, time1, time1)
        mmlc(718, 62, time1, time1, time1 * 10)
        keysendstring(i)
        sleep(time1 * 10)
        keypress(13)
        sleep(1000)
        mmlc(269, 62, time1, time1, time1)
        sleep(time1 * 5)
        mmlc(120, 1056, time1, time1, time1)
    end
    //    while(true)
    //         
    //    end
end
function xunhuannc()
    var i0 = btn_点击()
    messagebox(i0)
    for(var i = i0; i <= i0 + 25; i++)
        执行nc()
        // 点击下一页
        // 区域找色
        var x = -1, y = -1
        var ret = findcolor(1784, 879, 1825, 1030, "57BEC7-000000", 0.8, 0, x, y)
        //        点击
        mmlc(x, y, time1, time1, time1 * 5)
        mousemove(1874, 425)
        var cr2 = getcolor(1874, 425, 1)
        while(cr2 != "ececec")
            cr2 = getcolor(1874, 425, 1)
            sleep(200)
        end
        //        下载数据包
        mmlc(164, 1055, time1, time1, time1)
        mmlc(718, 62, time1, time1, time1 * 10)
        keysendstring(i)
        sleep(time1 * 10)
        keypress(13)
        sleep(1000)
        mmlc(269, 62, time1, time1, time1)
        sleep(time1 * 5)
        mmlc(120, 1056, time1, time1, time1)
    end
    //    while(true)
    //         
    //    end
end
//启动_热键操作
function 启动_热键()
    //    线程ID = threadbegin("xunhuancz", "")
    线程ID = threadbegin("xunhuannc", "")
end
//终止热键操作
function 终止_热键()
    threadclose(线程ID)
end
//鼠标移动点击
function mmlc(x, y, t1, t2, t3)
    sleep(t1)
    mousemove(x, y)
    sleep(t2)
    mouseleftclick(1)
    sleep(t3)
end
function btn_点击()
    //这里添加你要执行的代码
    var a = editgettext("text")
    return a
end