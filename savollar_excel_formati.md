# Savollar Excel Formati

Workbook ichida 2 xil sheet bo'ladi:

- Har bir yo'nalish uchun alohida savollar sheeti: `AI Junior`, `Python Backend`, `Data Science` kabi.
- Umumiy sozlama sheeti: `Yo'nalish sozlamalari`.

Sheet nomi ilovada yo'nalish nomi sifatida chiqadi.

## Yo'nalish Sozlamalari

`Yo'nalish sozlamalari` sheetida har bir yo'nalishning jami bali va o'tish foizi ko'rsatiladi.

| Ustun nomi | Izoh |
| --- | --- |
| `Yo'nalish` | Savollar sheet nomi bilan bir xil bo'lishi kerak |
| `Jami ball` | Shu yo'nalishdagi savollar `Maks ball` yig'indisi |
| `O'tish foizi` | O'tish uchun kerak bo'lgan minimal foiz, masalan `60` |
| `Izoh` | Ixtiyoriy izoh |

Ilova kandidatning foizini shunday hisoblaydi:

```text
Foiz = Olingan ball / Jami maksimal ball * 100
```

Keyin `Foiz >= O'tish foizi` bo'lsa natija `O'tdi`, aks holda `O'tmadi` bo'ladi.

## Savollar Sheetidagi Majburiy Ustunlar

| Ustun nomi | Izoh |
| --- | --- |
| `No` | Savol raqami |
| `Savol turi` | `test` yoki `yozma` |
| `Savol` | Kandidatga beriladigan savol matni |
| `To'g'ri javob` | Test savol uchun `A`, `B`, `C`, yoki `D` |
| `Maks ball` | Savol maksimal bali, masalan `1`, `2`, `5` |

## Test Savol Ustunlari

Test savolda quyidagilar to'ldiriladi:

| Ustun nomi | Izoh |
| --- | --- |
| `A variant` | A javob varianti |
| `B variant` | B javob varianti |
| `C variant` | C javob varianti |
| `D variant` | D javob varianti |
| `To'g'ri javob` | To'g'ri variant harfi: `A`, `B`, `C`, yoki `D` |

Test savolda `Kutilgan javob/Rubrika` bo'sh qolishi mumkin.

## Yozma Savol Ustunlari

Yozma savolda quyidagilar to'ldiriladi:

| Ustun nomi | Izoh |
| --- | --- |
| `Savol turi` | `yozma` |
| `Savol` | Yozma javob talab qiladigan savol |
| `Kutilgan javob/Rubrika` | AI baholash uchun mezon, kutilgan javob, kalit tushunchalar |
| `Maks ball` | Masalan `2`, `5`, `10` |

Yozma savolda `A variant`, `B variant`, `C variant`, `D variant`, `To'g'ri javob` bo'sh qoladi.

## Ixtiyoriy Ustunlar

| Ustun nomi | Izoh |
| --- | --- |
| `Qiyinlik darajasi` | `Oson`, `O'rta`, `Qiyin` |
| `Mavzu` | Savol mavzusi, masalan `ML asoslari`, `Python`, `SQL` |

## To'liq Savollar Ustun Tartibi

```text
No
Savol turi
Savol
A variant
B variant
C variant
D variant
To'g'ri javob
Kutilgan javob/Rubrika
Maks ball
Qiyinlik darajasi
Mavzu
```

## Namuna Qatorlar

| No | Savol turi | Savol | A variant | B variant | C variant | D variant | To'g'ri javob | Kutilgan javob/Rubrika | Maks ball | Qiyinlik darajasi | Mavzu |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | test | Python'da list nima? | O'zgarmas kolleksiya | Tartibli va o'zgaruvchan kolleksiya | Faqat matn turi | Funksiya turi | B |  | 1 | Oson | Python |
| 2 | yozma | Overfitting nima va uni qanday kamaytirish mumkin? |  |  |  |  |  | Overfitting model train dataga haddan tashqari moslashib, test datada yomon ishlashidir. Yechimlar: regularization, dropout, ko'proq data, early stopping, cross-validation. | 5 | O'rta | Model baholash |

## Natija Excelga Yoziladigan Asosiy Ustunlar

`candidate_results.xlsx` faylida quyidagilar saqlanadi:

- `Yo'nalish jami ball`
- `Jami maksimal ball`
- `Olingan ball`
- `Foiz`
- `O'tish foizi`
- `O'tish holati`
- `Baho`

## Muhim Qoidalar

- Har bir yo'nalish alohida sheet bo'lishi kerak.
- `Yo'nalish sozlamalari` sheetidagi `Yo'nalish` qiymati savollar sheet nomi bilan bir xil yozilsin.
- `Savol turi` faqat `test` yoki `yozma` bo'lsin.
- Test savolda `To'g'ri javob` faqat bitta harf bo'lsin: `A`, `B`, `C`, yoki `D`.
- `Maks ball`, `Jami ball`, `O'tish foizi` raqam bo'lishi kerak.
- Yozma savolda `Kutilgan javob/Rubrika` qanchalik aniq bo'lsa, AI baholash shunchalik yaxshi ishlaydi.
