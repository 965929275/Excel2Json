针对二维表格：

1. 确定json数据结构： 

   ```
   {
       "id0":{
           "key0":value,
           "key1":value,
           "key2":value,
           "key3":value,
           .
           .
           .
       },
       "id1":{
           "key0":value,
           "key1":value,
           "key2":value,
           "key3":value,
           .
           .
           .
       },
       "id2":{},
       "id3":{},
       .
       .
       .
   }
   ```

2. 我的算法：

   - 循环遍历第一行，得到所有的**key**，存在列表**keys**中备用。
   - 循环遍历第一列，得到所有的**id**，存在列表**ids**中备用。

   - 创建一个字典dict
   - 遍历ids，给每个id赋值dict_row
   - 创建dict_row
   - 遍历keys，给每个key赋值value，存到dict_row中
   - 取value的值：遍历rows，得到cell.value

3. 学习网上的算法：

   - 抓取所有sheet，可指定解析某一个sheet，或者循环解析所有sheet生成多个json文件。
   - 定义一个json对象**result**为转换的结果
   - 获取行号**id**作为每行内容的**id**
   - 遍历所有行，作为for循环的第一层
     - 定义一个json对象**line**作为每行内容的容器
     - 遍历所有列，获得每列的key和value
     - 将id与line组合后放入reslut中
   - json.dump()保存
