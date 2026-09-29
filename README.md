## 1. Best solution: LAMBDA + Name Manager

After creating the function once, you can simply write:

```excel
=NumToWordsIN(A2)
```

or:

```excel
=RupeesInWords(A2)
```

This is much cleaner than putting a giant formula in every cell. Microsoft specifically recommends moving reusable LAMBDA functions into the Name Manager.

### Step 1: Create `NumToWordsIN`

Go to **Formulas → Name Manager → New**.

**Name:** `NumToWordsIN`

**Refers to:**

```excel
=LAMBDA(n,
LET(
	x,INT(ABS(n)),
	ones,{"","One","Two","Three","Four","Five","Six","Seven","Eight","Nine"},
	teens,{"Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen","Seventeen","Eighteen","Nineteen"},
	tens,{"","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"},
	two,LAMBDA(k,
		IF(k<10,
			INDEX(ones,1,k+1),
			IF(k<20,
				INDEX(teens,1,k-9),
				INDEX(tens,1,INT(k/10)+1)&
				IF(MOD(k,10)>0," "&INDEX(ones,1,MOD(k,10)+1),"")
			)
		)
	),
	three,LAMBDA(k,
		IF(k<100,
			two(k),
			INDEX(ones,1,INT(k/100)+1)&" Hundred"&
			IF(MOD(k,100)>0," and "&two(MOD(k,100)),"")
		)
	),
	result,
		IF(x=0,
			"Zero",
			IF(x<1000,
				three(x),
				IF(x<100000,
					three(QUOTIENT(x,1000))&" Thousand"&
					IF(MOD(x,1000)>0," "&three(MOD(x,1000)),""),
					IF(x<10000000,
						three(QUOTIENT(x,100000))&" Lakh"&
						IF(MOD(x,100000)>0," "&NumToWordsIN(MOD(x,100000)),""),
						NumToWordsIN(QUOTIENT(x,10000000))&" Crore"&
						IF(MOD(x,10000000)>0," "&NumToWordsIN(MOD(x,10000000)),"")
					)
				)
			)
		),
	IF(n<0,"Minus "&result,result)
))
```

Click **OK**.

This uses recursive LAMBDA calls for the higher Indian-number groups. Excel supports recursive LAMBDA functions, although excessive recursion can produce a `#NUM!` error.

### Step 2: Use it

Suppose `A2` contains `123456789`. Enter:

```excel
=NumToWordsIN(A2)
```

Result: **Twelve Crore Thirty Four Lakh Fifty Six Thousand Seven Hundred and Eighty Nine**

| Number | Result |
| --- | --- |
| 0 | Zero |
| 21 | Twenty One |
| 100 | One Hundred |
| 1,234 | One Thousand Two Hundred and Thirty Four |
| 1,00,000 | One Lakh |
| 12,34,567 | Twelve Lakh Thirty Four Thousand Five Hundred and Sixty Seven |
| 1,00,00,000 | One Crore |
| 12,34,56,789 | Twelve Crore Thirty Four Lakh Fifty Six Thousand Seven Hundred and Eighty Nine |
| 1,00,00,00,000 | One Hundred Crore |

The function also handles negative numbers. For example, `-25000` becomes **Minus Twenty Five Thousand**.

## 2. For invoices: Rupees + Paise

For Indian invoices, quotations, receipts, payment vouchers, and similar documents, a dedicated currency function is more useful.

Create another name via **Formulas → Name Manager → New**.

**Name:** `RupeesInWords`

**Refers to:**

```excel
=LAMBDA(amount,
LET(
	totalPaise,ROUND(ABS(amount)*100,0),
	rupees,QUOTIENT(totalPaise,100),
	paise,MOD(totalPaise,100),
	IF(amount<0,"Minus ","")&
	"Rupees "&
	NumToWordsIN(rupees)&
	IF(paise=0,
		"",
		IF(paise=1,
			" and One Paisa",
			" and "&NumToWordsIN(paise)&" Paise"
		)
	)&
	" Only"
))
```

Then use:

```excel
=RupeesInWords(A2)
```

| Amount | Result |
| --- | --- |
| 1234567.50 | Rupees Twelve Lakh Thirty Four Thousand Five Hundred and Sixty Seven and Fifty Paise Only |
| 125000 | Rupees One Lakh Twenty Five Thousand Only |
| 0 | Rupees Zero Only |

## 3. Why LAMBDA is the best method

| Method | VBA-free | Works directly in cells | Reusable | Internet required | Assessment |
| --- | --- | --- | --- | --- | --- |
| LAMBDA | Yes | Yes | Yes | No | Best for modern Excel |
| Giant worksheet formula | Yes | Yes | No | No | Works, but difficult to maintain |
| Power Query | Yes | No (not live-cell style) | Yes | No | Good for batch data transformation |
| Third-party add-in | Yes | Yes | Yes | Usually no | Convenient, but adds dependency |
| XML/web-service tricks | Yes | Yes | Sometimes | Yes | Avoid for important work |
| VBA UDF | No | Yes | Yes | No | Not suitable when VBA is prohibited |
