# -
实现五年数据的数据分析
re = "REPORT";
vers = "V100R001C10B001_TA";
sed -i "/^\[$re\]/,/^$/ s/reportVerison=.*$/reportVerison=$vers/" D:/cdn/REPORT-master/install/conf/agent.ini
