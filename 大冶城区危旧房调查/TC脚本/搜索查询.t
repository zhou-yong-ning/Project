var 线程ID
var time1 = 50
//从这里开始执行
function 执行()
    for(var i = 0; i < 30; i++)
        sleep(time1 * 2)
        复制()
        sleep(time1 * 5)
        //任务栏图标三
        mmlc(122, 1057, time1 * 5, time1, time1 * 5)
        //点击搜索框
        mmlc(973, 165, 0, time1, time1)
        全选()
        sleep(time1)
        粘贴()
        sleep(time1)
        //点击查询
        mmlc(1075, 165, 0, time1 * 3, 6000)
        //查询结果反馈
        var ret = -1, x = -1, y = -1
        while(ret < 0)
            ret = findcolor(323, 445, 865, 814, "00ffff-000000", 0.9, 0, x, y)
            sleep(200)
        end
        //任务栏图标四
        mmlc(208, 1057, 0, time1, time1 * 2)
        //    键盘下键
        keypress(40)
    end
end
function xunhuan()
    var a = editgettext("编辑框0")
    messagebox(a)
    for(var i = a; i <= 100; i++)
        执行()
        //下载数据包
        mmlc(164, 1055, time1, time1, time1)
        mmlc(718, 62, time1 * 5, time1, time1 * 10)
        keysendstring(i)
        sleep(time1 * 10)
        keypress(13)
        sleep(1000)
        mmlc(269, 62, time1 * 2, time1 * 3, time1 * 4)
        sleep(time1 * 5)
        mmlc(208, 1056, time1, time1, time1 * 5)
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