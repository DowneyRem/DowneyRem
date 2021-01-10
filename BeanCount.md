# BeanCount | 复式记账


### 安装 [BeanCount]() 与 [Fava]()
Beancount 是复式记账的核心 命令行  
Fava 则是Beancount 的图形界面
```
pip install BeanCount fava
```

### 安装好了，如何记账？
尽管 Beancount 是通过命令行操作的，还是可以通过图形界面来编写账本
#### 编写 main.bean 
新建TXT，更改名字为 `main.bean` 而非 `main.bean.txt` 
```
; 设定账本标题
option "title" "我的账本"
; 设定账本主货币，人民币
option "operating_currency" "CNY"
; Fave的设定
1900-01-01 custom "fava-option" "language" "zh_cn" "default-file" "2021.bean"
; 所有使用中的账户都写在 accounts.bean
include "accounts.bean"
; 每年的账本
include "2020.bean"
```

#### 编写 accounts.bean 大小账户
你需要将收支资债分为 5个账户大类，然后再进一步细分
```
; 资产 Assets —— 现金、银行存款、有价证券等；
; 负债 Liabilities —— 信用卡、房贷、车贷等；
; 收入 Income —— 工资、奖金等；
; 费用 Expenses —— 外出就餐、购物、旅行等；
; 权益 Equity —— 用于「存放」某个时间段开始前已有的豆子，下文详述。
```

#### 编写 2020.bean 具体的账本


### 参考  
#### BY byvoid [博文](https://byvoid.com/zhs/) [OpenCC项目](https://github.com/BYVoid/OpenCC)作者   
1. [Beancount复式记账（2）：为什么](https://byvoid.com/zhs/blog/beancount-bookkeeping-1/)  
2. [Beancount复式记账（3）：为什么](https://byvoid.com/zhs/blog/beancount-bookkeeping-2/)  
3. [Beancount复式记账（3）：为什么](https://byvoid.com/zhs/blog/beancount-bookkeeping-3/)  
4. [Beancount复式记账（4）：为什么](https://byvoid.com/zhs/blog/beancount-bookkeeping-4/)  

[Beancount复式记账（4）：为什么](https://byvoid.com/zhs/blog/beancount-bookkeeping-4/)  

#### By Lyric  
0. [复式记账基础与 BEANCOUNT](https://gitpress.io/c/beancount-tutorial/beancount-tutorial-0)  
1. [握着你的手写最简单的 BEANCOUNT 账本](https://gitpress.io/c/beancount-tutorial/beancount-tutorial-1)  
2. [导入浦发银行信用卡账单](https://gitpress.io/c/beancount-tutorial/beancount-tutorial-2)  
3. [BEANCOUNT 速查表](https://gitpress.io/c/beancount-tutorial/beancount-tutorial-3)  
4. [BEANCOUNT 速查表](https://gitpress.io/c/beancount-tutorial/beancount-cheat-sheet)

#### 单篇

[Beancount —— 命令行复式簿记](https://wzyboy.im/post/1063.html) by wzyboy  
[使用 BEANCOUNT 自动记账](https://lyric.im/beancount) by  lyric 安利自上文





