var 线程ID
//从这里开始执行
function 执行()
    //从这里开始你的代码
    var time1 = 40
    sleep(time1 * 2)
    复制()
    sleep(time1)
    //任务栏图标三
    mmlc(122, 1057, 0, time1, time1)
    //点击查询
    var cr1 = getcolor(35, 231, 1)
    while(cr1 != "008a96")
        cr1 = getcolor(35, 231, 1)
        sleep(200)
    end
    sleep(time1 * 5)
    //点击查询
    mmlc(1292, 171, time1 * 5, time1, time1)
    //弹出界面操作
    cr1 = getcolor(1484, 320, 1)
    while(cr1 != "f8f8f8")
        cr1 = getcolor(1484, 320, 1)
        sleep(200)
    end
    sleep(time1 * 2)
    //点击搜索框
    mmlc(593, 644, 0, time1, time1)
    全选()
    sleep(time1)
    粘贴()
    sleep(time1)
    //点击查询
    mmlc(863, 891, 0, time1 * 3, time1 * 3)
    //查询结果点击
    cr1 = getcolor(30, 362, 1)
    while(cr1 != "ffffff")
        cr1 = getcolor(30, 362, 1)
        sleep(200)
    end  
    mmlc(125, 343, time1 * 3, time1 * 3, 0)
    cr1 = getcolor(957, 993, 1)
    while(cr1 != "068996")
        cr1 = getcolor(957, 993, 1)
        sleep(200)
    end 
    sleep(time1 * 10)
    mmlc(433, 273, time1, time1, time1)
    sleep(time1 * 10)
    //点击返回
    mmlc(1849, 189, time1 * 3, time1 * 3, time1 * 3)
    //任务栏图标四
    mmlc(170, 1057, 0, time1, time1 * 2)
    sleep(time1 * 2)
    keypress(40)
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