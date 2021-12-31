import cx_Oracle


def oracle_dao(**kwargs):
    """
    Oracle数据库连接
    :param kwargs: 数据库连接地址,格式[用户名]/[密码]@[host]:[port]/[SERVICE_NAME]（可选）
    ，默认path="XIPC_ZNYP/XIPC_ZNYP@172.16.2.89:1521/ywptdata"
    :return:
    """
    path = "XIPC_ZNYP/XIPC_ZNYP@172.16.2.89:1521/ywptdata"
    if kwargs:
        path = kwargs["path"]
    db = cx_Oracle.connect(path)
    return db
