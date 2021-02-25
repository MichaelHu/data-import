-- ['项目', '战区', '省份', '地市', '项目经理', '预计收入', '预计合同年', '状态', '所属部门', '类型', '预计人月', '额外成本（万）', '其他投入（人时）', '2020其他投入（人时）', '190601']
create table t_projects (
    id integer primary key autoincrement not null,
    name text not null,
    zhanqu text not null,
    shengfen text not null,
    dishi text not null,
    xiangmujingli text not null,
    yujishouru numeric not null,
    yujihetongnian text not null,
    zhuangtai text not null,
    suoshubumen text not null,
    leixing text not null,
    yujirenyue numeric not null,
    ewaichengben numeric not null,
    qitatouru numeric not null,
    qitatouru2020 numeric not null,
    qitatouru190601 numeric not null
);
