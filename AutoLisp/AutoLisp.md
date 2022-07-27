# AutoLisp

## 递归函数

```lisp
# 求阶乘
(defun jc(n)
 	(if (= n 0)
  		1
  		(* n (jc(- n 1)))
	  )
       )
(defun c:qjc()
	(setq n(getint "输入整数："))
  	(jc n)
  )
```

## 调用cad命令函数

```lisp
# "command":调用工具画圆命令
(defun c:hy()
  	(setq pc (getpoint "请输入圆心："))
  	(setq r (getint "请输入半径："))
  	(command "circle" pc r)
 )
```

```lisp
# (setvar "cmdecho" 0)用来去除系统提示
(defun c:hy()
    (setvar "cmdecho" 0)
  	(setq pc (getpoint "请输入圆心："))
  	(setq r (getdist pc "请输入半径："))
  	(command "circle" pc r)
 )
```

```lisp
# 三点画圆
(defun c:hy()
  	(setvar "cmdecho" 0)
  	(setq p1 (getpoint "请输入点一："))
  	(setq p2 (getpoint "请输入点二："))
  	(setq p3 (getpoint "请输入点三："))
  	(command "circle" "3p" p1 p2 p3)
 )
```

```lisp
# ""代表结束后敲击回车建
(defun c:x()
	(command "line" '(1 2) '(5 6) "")
  )
```

