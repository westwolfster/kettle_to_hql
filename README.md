【星海hql装修工具功能】
将能力开放门户导出的ETL流程xml文件，根据执行顺序（jobhops）还原为原始的hql脚本。导出的xml文件应包含主流程xml文件和所包含的若干子流程xml文件，全部打包成zip文件，作为装修工具的输入。
装修工具将输入的zip包，按照星海数据中台的脚本编写规则，将xml流程定义文件转换为hql脚本文件。输出的hql脚本文件以主流程为单位，一个主流程输出一个hql文件，hql文件名以主流程名称命名。
【星海hql装修工具处理规则】
一、xml流程定义文件中的非核心组件处理
1、时间参数组件，直接阉掉，不做任何处理；
2、数据到达等待检查组件(等待条件SQL)、（MYSQL 元数据库）部分，不处理SQL脚本，在该组件出现的位置，以注释的形式体现；
3、MYSQL日志写入组件，不做SQL处理，提取mysql插入语句中的table_name字段内容，以注释的形式体现；
4、MYSQL 下游采集触发组件，不做SQL处理，提取mysql插入语句中的table_name字段内容，以注释的形式体现。
二、xml流程定义文件中的核心HIVE业务SQL组件处理
1、阉掉所有set语句群；
2、所有的表名，一律改为小写；
3、SQL脚本中创建的临时表，表名前缀统一改为"temp_"。
4、SQL脚本中没有模式名称的表名，以"模式名称_dev"作为新的模式名，例如将"TMP_CDMA_CHARGE_TJ0_${month_no}"，改为"模式名称_dev.temp_cdma_charge_tj0_${month_no}"。
有模式名的表名，将模式名的后面加上"_dev"作为新的模式名。
5、对原脚本中的注释，不做任何处理，包括注释的格式、出现的位置和注释的内容。
6、对原脚本中出现的"ifnull"函数，统一改为"coalesce"函数。
7、在流程输出的SQL脚本末尾，如果没有删除本流程建立的临时表语句群，会新增删除所有本流程建立的临时表语句群。
【星海hql装修工具使用方法】
python kettle_to_hql_general.py --zip ${xml文件包所在路径}\${xml文件包名称}.zip --out ${输出文件夹所在路径} --schema ${模式名称}
例如：C:\Users\westw\Downloads>python kettle_to_hql_general.py --zip C:\Users\westw\Downloads\20260726.zip --out C:\Users\westw\Downloads\output_20260726 --schema wt_day 
