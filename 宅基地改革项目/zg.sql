-- 连接QLR与QLRGL表
SELECT QLR.QLRMC 户主,GL.QLRMC 家庭成员,GL.ZJHM 证件号码 FROM DJ_QLR QLR INNER JOIN DJ_QLRGL GL ON QLR.QLRID=GL.QLRID;

SELECT ZDQLR.QLRMC,ZDQLR.ZDTYBM FROM ZD_QSDC ZDQLR;

SELECT ZDTYBM,COUNT(*) AS 数量 FROM ZD_QSDC GROUP BY ZDTYBM HAVING 数量 = 1;


-- 查询宗地号对应的家庭成员
-- CREATE TABLE FDYT AS
SELECT A.ZDTYBM AS 宗地号,
       DJ_QLR.QLRMC AS 权利人名称,
       DJ_QLR.ZJHM AS 证件号码,
       DJ_QLR.GX AS 与户主关系,
       leftstr(A.ZDTYBM,12) AS 村,
       row_number() over () AS ID
FROM DJ_QLR INNER JOIN (SELECT ZD_QSDC.ZDTYBM,DJ_QLRGL.QLRID AS 权利人ID
                        FROM ZD_QSDC INNER JOIN DJ_QLRGL ON TSTYBM = YWBM) AS A ON DJ_QLR.QLRID = A.权利人ID;

-- 查询宗地号对应的家庭成员和区划(并且添加行号为ID主键)
SELECT 宗地号, 权利人名称, 证件号码, 与户主关系, 区域代码, QYMC AS 区域名称,ID
FROM (SELECT A.ZDTYBM AS 宗地号,
       DJ_QLR.QLRMC AS 权利人名称,
       DJ_QLR.ZJHM AS 证件号码,
       DJ_QLR.GX AS 与户主关系,
       leftstr(A.ZDTYBM,12) AS 区域代码,
       row_number() over () AS ID
FROM DJ_QLR INNER JOIN (SELECT ZD_QSDC.ZDTYBM,DJ_QLRGL.QLRID AS 权利人ID
                        FROM ZD_QSDC INNER JOIN DJ_QLRGL ON TSTYBM = YWBM) AS A ON DJ_QLR.QLRID = A.权利人ID) AS FDYT
INNER JOIN QH ON FDYT.区域代码 = QH.QYDM;


SELECT 宗地号, 权利人名称, 证件号码, 与户主关系, 村代码, 区划, ID,
CASE WHEN instr(区划,'镇')>0 THEN substr(区划,1,instr(区划,'镇'))
    WHEN instr(区划,'乡') > 0 THEN substr(区划,1,instr(区划,'乡'))
    ELSE 区划 END AS 乡镇
FROM FDYTCY WHERE 区划 LIKE '%茗山%';

-- 宅基地根据房地一体数据查询
SELECT 宗地号, 权利人名称, 证件号码, 与户主关系, 村代码, 区划, id FROM FDYTCY
WHERE 权利人名称 = '陈绪安';
SELECT 宗地号, 权利人名称, 证件号码, 与户主关系, 村代码, 区划, id FROM FDYTCY
WHERE 证件号码 = '420281198510015127';
SELECT 宗地号, 权利人名称, 证件号码, 与户主关系, 村代码, 区划, ID FROM FDYTCY
WHERE 宗地号 IN (SELECT 宗地号 FROM FDYTCY
WHERE 证件号码 = '420281199410252453');
SELECT 宗地号, 权利人名称, 证件号码, 与户主关系, 村代码, 区划, id FROM FDYTCY
WHERE 宗地号 = '420281198212072430';
-- 宅基地根据农村集体经济组织成员查询
SELECT 成员户编码, 组名, 姓名, 性别, 民族, 身份证号码, 与户主关系, 户籍地址, 户主联系电话 FROM NCJTJJZZ
WHERE 姓名 = '石进';
SELECT 成员户编码, 组名, 姓名, 性别, 民族, 身份证号码, 与户主关系, 户籍地址, 户主联系电话 FROM NCJTJJZZ
WHERE 成员户编码 IN (SELECT 成员户编码 FROM NCJTJJZZ
WHERE 身份证号码 = '420221196601240076');


-- 标准地名信息查询
SELECT DZ, XQMC, XZQH, XM, ZJHM FROM BZDMXX
WHERE ZJHM = '420221197705082839';
-- 按村查询
SELECT 宗地号, 权利人名称, 证件号码, 与户主关系, 村代码, 区划, id FROM FDYTCY
WHERE  区划 LIKE '%桂树%';
-- 建立索引
CREATE INDEX SY ON FDYTCY(ID);
-- 联合房地一体数据库与农经权
CREATE TABLE NFNJQ AS
SELECT 宗地号 AS 家庭户编码, 权利人名称 AS 成员, 证件号码, 与户主关系, 区划 FROM FDYTCY
UNION
SELECT 成员户编码 AS 家庭户编码,姓名 AS 成员,身份证号码 AS 证件号码, 与户主关系, 户籍地址 AS 区划 FROM NCJTJJZZ;

UPDATE NFNJQ SET 数据来源 = CASE WHEN length(NFNJQ.家庭户编码) > 6 THEN '房地一体' ELSE '农经权' END;

-- 新增删除列
ALTER TABLE NFNJQ DROP COLUMN ID;
ALTER TABLE NFNJQ ADD COLUMN 数据来源 TEXT;
ALTER TABLE NFNJQ ADD COLUMN ID TEXT;
-- 创建表
CREATE TABLE NJQNF AS
SELECT *,row_number() over () AS ID FROM NFNJQ;

-- 创建索引
CREATE INDEX ID1 ON NFNJQ(家庭户编码);
DROP INDEX ID1;
-- 查询联合表
SELECT 家庭户编码, 成员, 证件号码, 与户主关系, 区划,数据来源 FROM NFNJQ
WHERE 家庭户编码 IN (SELECT 家庭户编码 FROM NFNJQ
WHERE 证件号码 = '420221195909222830');

SELECT 家庭户编码, 成员, 证件号码, 与户主关系, 区划,数据来源 FROM NFNJQ
WHERE 家庭户编码 IN (SELECT 家庭户编码 FROM NFNJQ
WHERE 成员 = '张伟');

SELECT 家庭户编码, 成员, 证件号码, 与户主关系, 区划,数据来源 FROM NFNJQ
WHERE 家庭户编码='420281007014JC11011';

SELECT NFNJQ.家庭户编码, NFNJQ.成员, NFNJQ.证件号码, NFNJQ.与户主关系, NFNJQ.区划, NFNJQ.数据来源 FROM NFNJQ
WHERE 证件号码 = '420221196506086947';
SELECT NFNJQ.家庭户编码, NFNJQ.成员, NFNJQ.证件号码, NFNJQ.与户主关系, NFNJQ.区划, NFNJQ.数据来源 FROM NFNJQ
WHERE 成员 = '陈绪安';

-- 数据表处理
SELECT * FROM NFNJQ WHERE 区划 LIKE '%金山店%';

DELETE FROM NFNJQ WHERE 成员 = '姓名' AND 证件号码 = '身份证号码';
-- 分组顺编宗地号
SELECT *,row_number() over (partition by 家庭户编码) AS BM FROM NFNJQ;

UPDATE NFNJQ SET 证件号码 = replace(证件号码,'x','X');