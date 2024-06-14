# -
实现五年数据的数据分析

chmod 640 /etc/crontab
crontab -u root /etc/crontab
rm -f /etc/cron.daily/um_compress_log  > /dev/null 2>&1
rm -f /etc/cron.hourly/um_disk_check.sh  > /dev/null 2>&1 为什么这段代码crontab -u root /etc/crontab，此命令会将使用crontab -e命令添加的定时任务清理掉，因此导致局点自己添加的定时任务丢失
