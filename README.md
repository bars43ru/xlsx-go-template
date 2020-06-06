# Заполнение документов на основе шаблона xlsx 

### Идея и часть реализации взята из проекта https://github.com/ivahaev/go-xlsx-templater

Были выполнены следующие доработки:
* Заполнение шаблона из struct а не map[] – данная доработка обеспечила возможность вызова функций (`реализуемые  пользовательским типом`) в шаблоне и использование возвращаемого значения для заполнения документа
* Изменение высоты строк на основе данных для заполнения
* Не теряется форматирование числовых значений в заполненном документе
