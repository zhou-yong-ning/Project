var 线程ID
//从这里开始执行
function 执行()
    //从这里开始你的代码
    var time1 = 10
    sleep(time1 * 2)
    复制()
    sleep(time1)
    //任务栏图标三
    mmlc(122, 1057, 0, time1, time1)
    //点击查询
    var cr1 = getcolor(1898, 98, 1)
    while(cr1 != "2e4777")
        cr1 = getcolor(1898, 98, 1)
        sleep(200)
    end
    //点击查询
    mmlc(214, 249, time1 * 3, time1, time1)
    //弹出界面操作
    cr1 = getcolor(1896, 230, 1)
    while(cr1 != "67c23a")
        cr1 = getcolor(1896, 230, 1)
        sleep(200)
    end
    sleep(time1 * 2)
    //点击搜索框
    mmlc(457, 228, 0, time1, time1)
    全选()
    sleep(time1)
    粘贴()
    sleep(time1)
    //点击查询
    mmlc(1443, 234, 0, time1 * 3, 0)
    //查询结果点击
    mousemove(1893, 341)
    cr1 = getcolor(1699, 340, 1)
    while(cr1 != "f5f7fa")
        cr1 = getcolor(1699, 340, 1)
        sleep(200)
    end
    sleep(time1)
    if(getcolor(224, 395, 1) == "e6e6e6")
    else
        mouseleftclick(1)
    end
    //等待弹出窗口
    cr1 = getcolor(1874, 221, 1)
    while(cr1 != "409eff")
        cr1 = getcolor(1874, 221, 1)
        sleep(200)
    end
    //查看照片
    mmlc(1201, 288, time1 * 30, time1, 0)
    //点击照片
    mmlc(1127, 430, time1 * 20, time1, 0)
    //按键控制
    var aj = keywait()
    var text
    while(aj)
        if(aj == 97)
            text = "1"
            break
        elseif(aj == 98)
            text = "2"
            break
        elseif(aj == 99)
            text = "3"
            break
            //        	  左键
        elseif(aj == 37)
            mmlc(1094, 662, 0, time1, 0)
            //        	  右键
        elseif(aj == 39)
            mmlc(1860, 662, 0, time1, 0)
            //            上
        elseif(aj == 38)
            mousemove(548, 617)
            sleep(time1)
            mouseleftdclick(1)
            //            下
        elseif(aj == 40)
            mousemove(548, 617)
            sleep(time1)
            mousewheelup(1)
            //            del
        elseif(aj == 46)
            mmlc(1536, 936, 0, time1, time1 * 5)
            keypress(13)
        end
        aj = keywait()
    end
    //任务栏图标四
    mmlc(170, 1057, 0, time1, time1 * 2)
    //    键盘右键赋值
    keypress(39)
    sleep(time1 * 5)
    keysendstring(text)
    keypress(40)
    //    键盘右键赋值第二格
    keypress(39)
    aj = keywait()
    while(true)
        if(aj == 97)
            text = "1"
            break
        elseif(aj == 98)
            text = "2"
            break
        elseif(aj == 99)
            text = "3"
            break
        elseif(aj == 100)
            text = "4"
            break
        elseif(aj == 101)
            text = "5"
            break
        elseif(aj == 96)
            text = "0"
            break
        end
    end
    //    键盘右键条件赋值第三格
    if(text == "1" || text == "3")        
        keypress(40)
        keypress(39)
        aj = keywait()
        while(aj != 13)
            aj = keywait()
            if(aj == 13)
                keypress(37, 3)
                break
            end
        end
    else
        keypress(40)
        keypress(37, 2)
        keypress(40)
    end
end
function xunhuan()
    while(true)
        执行()
    end
    //    for(var i = 0; i < 100; i++)
    //        执行()
    //    end
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