import os
import pandas as pd
import xlwt
import xlrd
from xlutils.copy import copy
import re
import tarfile
import gzip
from DelToExcel import DelToExcel
import zipfile
import unlzw3
from pathlib import Path
import paramiko
import requests
from RemoteOperation import RemoteOperation

def unTar123(fileName):
    """
       解压tar文件
       :param fname: 压缩文件名
       :param dirs: 解压后的存放路径
       :return: bool
    """
    try:
        t = tarfile.open(fileName)
        # t.extractall()
        names = t.getnames()
        dirs = fileName + "_files"
        # # 创建文件夹用来存放解压后的文件
        # if os.path.isdir(dirs):
        #     pass
        # else:
        #     os.mkdir(dirs)
        # for name in names:
        #     t.extract(name, dirs)
        # t.close()
        return True
    except Exception as e:
        print(e)
        return False


def un_zip(file_name):
    """unzip zip file"""
    zip_file = zipfile.ZipFile(file_name)
    if os.path.isdir(file_name + "_files"):
        pass
    else:
        os.mkdir(file_name + "_files")
    for names in zip_file.namelist():
        zip_file.extract(names, file_name + "_files/")
    zip_file.close()


def un_gz(file_name):
    """ungz zip file"""
    f_name = file_name.replace(".gz", "")
    # f_name = r"C:\Users\admin\Desktop\789\321.txt"
    # 获取文件的名称，去掉
    g_file = gzip.GzipFile(file_name)
    # 创建gzip对象
    open(f_name, "wb").write(g_file.read())
    # gzip对象用read()打开后，写入open()建立的文件里。
    g_file.close()
    # 关闭gzip对象


def untar(filename):
    tar = tarfile.open(filename)
    names = tar.getnames()
    # tar本身是将文件打包，解除打包会产生很多文件，因此需要建立文件夹存放
    if not os.path.isdir(filename + "_dir"):
        os.mkdir(filename + "_dir")
    for name in names:
        tar.extract(name, r"C:\Users\admin\Desktop\789")
    tar.close()

def uncompress(filesPath):
    fileslist = os.listdir(filesPath)
    for path in fileslist:
        # 只找压缩文件
        if 'tar.Z' not in path:
            filepath = filesPath + '\\' + path
            if os.path.isfile(filepath):
                unTar(filepath)


def decompress_pk(src, dst):
    """Decompress function using gzip & unlzw packages."""
    uc_func = unlzw3.unlzw if src.endswith('.Z') else gzip.decompress
    with open(src, 'rb') as sf, open(dst, 'wb') as df:
        buffer = sf.read()
        df.write(uc_func(buffer))


# 远程执行 CMD 命令， 并实时显示脚本执行情况
def ssh_Run_Cmd(host, username, password, cmd):
    """

    :param host:  主机 Ip
    :param username: 用户名 root
    :param password: 密码   Troila12
    :param cmd 想执行的命令
    :return:
    """
    ssh = paramiko.SSHClient()
    # 允许连接不在know_hosts文件中的主机
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    # 连接服务器
    ssh.connect(hostname=host, port=22, username=username, password=password)
    # 执行命令
    # stdin, stdout, stderr = ssh.exec_command('/usr/bin/Rscript /tandelindata/code.R')
    stdin, stdout, stderr = ssh.exec_command(cmd)
    # 获取命令结果
    # result = stdout.read().decode('utf-8')
    res = []  # 用于判断脚本是否执行完毕
    while len(res) < 10:
        result = stdout.readline().strip()
        if result is not None and len(result) != 0:
            # requests.request('post', 'url', data="result")
            print(result)
            res = []
        else:
            res.append(0)
    # 关闭连接
    ssh.close()
    # return result



#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
通过paramiko从远处服务器下载文件资源到本地
author: gxcuizy
time: 2018-08-01
"""

import paramiko
import os
from stat import S_ISDIR as isdir


def down_from_remote(sftp_obj, remote_dir_name, local_dir_name):
    """远程下载文件"""
    remote_file = sftp_obj.stat(remote_dir_name)
    if isdir(remote_file.st_mode):
        # 文件夹，不能直接下载，需要继续循环
        check_local_dir(local_dir_name)
        print('开始下载文件夹：' + remote_dir_name)
        for remote_file_name in sftp.listdir(remote_dir_name):
            sub_remote = os.path.join(remote_dir_name, remote_file_name)
            sub_remote = sub_remote.replace('\\', '/')
            sub_local = os.path.join(local_dir_name, remote_file_name)
            sub_local = sub_local.replace('\\', '/')
            down_from_remote(sftp_obj, sub_remote, sub_local)
    else:
        # 文件，直接下载
        print('开始下载文件：' + remote_dir_name)
        sftp.get(remote_dir_name, local_dir_name)


def check_local_dir(local_dir_name):
    """本地文件夹是否存在，不存在则创建"""
    if not os.path.exists(local_dir_name):
        os.makedirs(local_dir_name)


if __name__ == "__main__":
    """程序主入口"""
    # 服务器连接信息
    host_name = '81.68.187.65'
    user_name = 'root'
    password = 'Yjx123456'
    port = 22
    # 远程文件路径（需要绝对路径）
    remote_dir = '/home/test'
    # 本地文件存放路径（绝对路径或者相对路径都可以）
    local_dir = "C:/Users/admin/Desktop/123/"
    #
    aa = RemoteOperation(host_name, user_name, password)
    aa.download(remote_dir, local_dir)

    # 连接远程服务器
    # t = paramiko.Transport((host_name, port))
    # t.connect(username=user_name, password=password)
    # sftp = paramiko.SFTPClient.from_transport(t)
    #
    # # 远程文件开始下载
    # down_from_remote(sftp, remote_dir, local_dir)
    #
    # # 关闭连接
    # t.close()

# if __name__ == "__main__":
#     # delFilesLongPath = r'C:\Users\admin\Desktop\PBC\delFiles'
#     # colNameLongPath = r'C:\Users\admin\Desktop\PBC\ColumnName'
#     #
#     # targetPath = r'C:\Users\admin\Desktop\456'
#     # aa = DelToExcel()
#     # aa.delToExcels(delFilesLongPath, colNameLongPath)
#     # 服务器连接信息
#     host_name = '81.68.187.65'
#     user_name = 'root'
#     password = 'Yjx123456'
#     port = 22
#     # 远程文件路径（需要绝对路径）
#     remote_dir = '/data/nfs/zdlh/pdf/2018/07/31'
#     # 本地文件存放路径（绝对路径或者相对路径都可以）
#     local_dir = 'file_download/'
#
#     ssh_Run_Cmd(host_name, user_name, password, "tar -xvf /home/test/828_new-20210731.tar -C /home/test")
    # un_gz(r"C:\Users\admin\Desktop\789\tmp_odsfs\jwm\pbc\export_chaifen\data\20210731\828\E0001H233010001-CLTYCK-20210731_828.del.gz")