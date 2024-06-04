# -
实现五年数据的数据分析
modifySystemParameter()
{
    sysParameterFile="/etc/sysctl.conf"
    key="$1"
    value="$2"

    if [[ -f ${sysParameterFile} ]]; then
        havekey=`grep -c "${key}" ${sysParameterFile}`
        if [[ "${havekey}" -ne 1 ]]; then
            echo "Parameter ${key} don't exist, add Parameter ${key} to file  ${sysParameterFile}"
            echo "${key} = ${value}" >> ${sysParameterFile}
        elif [[ "${havekey}" -eq 1 ]]; then
            needModify=`grep -c "${key} = ${value}" ${sysParameterFile}`
            if [[ "${needModify}" -ne 1 ]]; then
                echo "${key} needModify is ${needModify}"
                echo "Parameter ${key} is ${value} need to modify"
                sed -i "s/${key}.*=.*/${key} = ${value}/g" ${sysParameterFile}
            elif [[ "${needModify}" -eq 1 ]]; then
                echo "Parameter ${key} is ${value} don't need to modify"
            fi
        fi
    else
        echo "file ${sysParameterFile} don't exist,please check it"
        return 1
    fi
}
