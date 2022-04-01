---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: В этом примере показано, как получить вложения из почтового ящика Exchange.
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:48:02 PM
---
# Надстройка Outlook: получение вложений с сервера Exchange Server

**Содержание**

* [Сводка](#summary)
* [Предварительные требования](#prerequisites)
* [Ключевые компоненты примера](#components)
* [Описание кода](#codedescription)
* [Сборка и отладка](#build)
* [Устранение неполадок](#troubleshooting)
* [Вопросы и комментарии](#questions)
* [Дополнительные ресурсы](#additional-resources)

<a name="summary"></a>
## Резюме
В этом примере показано, как получить вложения из почтового ящика Exchange.

<a name="prerequisites"></a>
## Предварительные требования ##

Для этого примера требуется следующее:  

  - Visual Studio 2013 с обновлением 5 или Visual Studio 2015.  
  - Компьютер с Exchange 2013 и по крайней мере одной учетной записью электронной почты или учетной записью Office 365. Вы можете [присоединиться к Программе разработчика Office 365 и получить бесплатную годовую подписку на Office 365](https://aka.ms/devprogramsignup).
  - Любой браузер, поддерживающий ECMAScript 5.1, HTML5 и CSS3, например Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 или более поздние версии этих браузеров.
  - Опыт программирования на JavaScript и работы с веб-службами.

<a name="components"></a>
## Ключевые компоненты примера
Пример решения содержит следующие файлы:

- AttachmentExampleManifest.xml: Файл манифеста для надстройки Outlook.
- AppRead\Home\Home.html: Пользовательский интерфейс HTML для почтовой надстройки для Outlook.
- AppRead\Home\Home.js: Файл JavaScript, который обрабатывает отправку информации о вложениях в удаленную службу вложений, включенную в этот образец.

Проект AttachmentService определяет службу REST с помощью API WCF. Проект содержит следующие файлы:

- Controllers\AttachmentServiceController.cs: Сервисный объект, который обеспечивает бизнес-логику для примера сервиса.
- Models\ServiceRequest: Объект, представляющий веб-запрос. Содержимое объекта создается из объекта запроса JSON, отправленного из вашей почтовой надстройки.
- Models\Attachment.cs: Служебный объект, помогающий десериализовать объект JSON, отправленный почтовой надстройкой.
- Models\AttachmentDetails.cs: Объект, который представляет детали каждого вложения. Он предоставляет объект .NET Framework, соответствующий объекту `сведений о вложениях` почтовых надстроек.
- Models\ServiceResponse: Объект, представляющий ответ от веб-службы. Содержимое объекта сериализуется в объект JSON при отправке обратно в почтовую надстройку.
- Web.config; Связывает образец службы с конечной точкой веб-сервера.



<a name="codedescription"></a>
##Описание кода

В этом примере показано, как получить вложения из веб-службы, поддерживающей вашу почтовую надстройку. Например, вы можете создать службу, которая загружает фотографии на сайт общего доступа, или службу, которая хранит документы в репозиторий. Служба получает вложения непосредственно с сервера Exchange и не требует от клиента выполнения дополнительной обработки, чтобы получить вложение и затем отправить его в службу.

Образец состоит из двух частей. Первая часть, почтовое приложение, запускается в почтовом клиенте. Почтовая надстройка отображается, когда сообщение или встреча является активным элементом. Когда вы нажимаете кнопку **Тестировать вложения**, надстройка электронной почты отправляет сведения о вложении веб-службе, которая обрабатывает запрос. Служба использует следующие шаги для обработки вложений:

- Отправляет запрос операции [GetAttachment](http://msdn.microsoft.com/library/aa494316(v=exchg.150).aspx) на сервер Exchange, на котором размещен почтовый ящик. Сервер отвечает отправкой вложения в сервис. В этом примере служба просто записывает XML с сервера для отслеживания вывода.
- Возвращает количество вложений, обработанных в почтовом приложении.



<a name="build"></a>
## Сборка и отладка ##
**Примечание**. Почтовая надстройка будет активирована для любого сообщения электронной почты в папке «Входящие» пользователя, которое имеет одно или несколько вложений. Вы можете упростить тестирование надстройки, отправив одно или несколько сообщений электронной почты в свою тестовую учетную запись перед запуском примера надстройки.

1. Откройте решение в Visual Studio
2. Щелкните правой кнопкой мыши решение в обозревателе решений. Выберите пункт**Назначить запускаемые проекты**. 
3. Выберите **Общие свойства** и выберите **Запуск проекта**.
4. Убедитесь, что для проекта **Действие** для **AttachmentExampleService** установлено значение **Пуск**.
5. Нажмите клавишу F5, чтобы собрать и развернуть пример надстройки.
6. Подключитесь к учетной записи Exchange, указав адрес электронной почты и пароль для сервера Exchange 2013.
7. Разрешить серверу настроить почтовую учетную запись.
8. Войдите в учетную запись электронной почты, введя имя учетной записи и пароль. 
9. Выберите сообщение в папке «Входящие».
10. Подождите, пока над сообщением появится панель надстроек.
11. На панели надстроек щелкните **AttachmentExample**.
12. Когда появится надстройка почты, нажмите кнопку **TestAttachments**, чтобы отправить запрос на сервер Exchange.
13. Сервер ответит количеством обработанных вложений для элемента. Это должно равняться количеству вложений, которые содержит элемент.

<a name="troubleshooting"></a>
## Устранение
неполадок Ниже перечислены распространенные ошибки, которые могут возникнуть при использовании Outlook Web App для проверки почтовой надстройки для Outlook:

- Панель надстроек не отображается при выборе сообщения. В этом случае перезапустите приложение, выбрав **Отладка - Остановить отладку** в окне Visual Studio, затем нажмите клавишу F5, чтобы перестроить и развернуть надстройку. 
- Изменения в коде JavaScript могут не учитываться при развертывании и запуске надстройки. Если изменения не отобраны, очистите кэш в веб-браузере, выбрав **инструменты — свойства браузера** и нажав кнопку **удалить...**. Удалите временные файлы Интернета, а затем перезапустите надстройку. 

<a name="questions"></a>
## Вопросы и комментарии ##

- Если у вас возникли проблемы с запуском этого примера, [сообщите о неполадке](https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments/issues).
- Вопросы о разработке надстроек Office в целом следует размещать в [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Обязательно помечайте свои вопросы и комментарии тегом [office-addins].


<a name="additional-resources"></a>
## Дополнительные ресурсы ##

- [Дополнительные примеры надстроек](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Web API: Официальный сайт Microsoft ASP.NET](http://www.asp.net/web-api)
- [ Получить вложения с сервера Exchange ](http://msdn.microsoft.com/library/dn148008.aspx)

## Авторские права
(c) Корпорация Майкрософт (Microsoft Corporation), 2015. Все права защищены.


Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/). Дополнительные сведения см. в разделе [часто задаваемых вопросов о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).