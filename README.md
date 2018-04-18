# projph
用于体育考试编排考号，导出各种统计表、登记表
生成准考证、条形码等。

2019 年计划要修改StudPh中一个字段（查看其中注释）

从正反两个方面查数据问题：
select * from itemselect where jump_option+rope_option+globe_option+bend_
option > 0 and signid in (select signid from freeexam);

select * from itemselect where jump_option+rope_option+globe_option+bend_
option = 0 and signid not in (select signid from freeexam);