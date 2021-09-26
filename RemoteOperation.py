import paramiko
import requests
import paramiko
import os
from stat import S_ISDIR as isdir


class RemoteOperation:
    def __init__(self, host, username, password):
        self.host = host
        self.username = username
        self.password = password

    def ssh_Run_Cmd(self, cmd):
        """
        远程执行 CMD 命令， 并实时显示脚本执行情况
        :param cmd 想执行的命令
        :return:
        """
        ssh = paramiko.SSHClient()
        # 允许连接不在know_hosts文件中的主机
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        # 连接服务器
        ssh.connect(hostname=self.host, port=22, username=self.username, password=self.password)
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

    def down_from_remote(self, sftp_obj, remote_dir_name, local_dir_name):
        """远程下载文件"""
        remote_file = sftp_obj.stat(remote_dir_name)
        if isdir(remote_file.st_mode):
            # 文件夹，不能直接下载，需要继续循环
            self.check_local_dir(local_dir_name)
            print('开始下载文件夹：' + remote_dir_name)
            for remote_file_name in sftp_obj.listdir(remote_dir_name):
                sub_remote = os.path.join(remote_dir_name, remote_file_name)
                sub_remote = sub_remote.replace('\\', '/')
                sub_local = os.path.join(local_dir_name, remote_file_name)
                sub_local = sub_local.replace('\\', '/')
                # 递归
                self.down_from_remote(sftp_obj, sub_remote, sub_local)
        else:
            # 文件，直接下载
            print('开始下载文件：' + remote_dir_name)
            sftp_obj.get(remote_dir_name, local_dir_name)

    def check_local_dir(self, local_dir_name):
        """本地文件夹是否存在，不存在则创建"""
        if not os.path.exists(local_dir_name):
            os.makedirs(local_dir_name)

    def download(self, remote_dir, local_dir):
        """
        从远程服务器批量下载文件
        :param remote_dir:
        :param local_dir:
        :return:
        """
        # 服务器连接信息
        host_name = self.host
        user_name = self.username
        password = self.password
        port = 22

        # 连接远程服务器
        t = paramiko.Transport((host_name, port))
        t.connect(username=user_name, password=password)
        sftp = paramiko.SFTPClient.from_transport(t)

        # 远程文件开始下载
        self.down_from_remote(sftp, remote_dir, local_dir)

        # 关闭连接
        t.close()
