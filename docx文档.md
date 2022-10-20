# Docx文档操作

## 一、常用类、接口

常用类

```java
org.apache.poi.xwpf.usermodel.XWPFDocument		// 用于处理.docx文件的高级类。
org.apache.poi.xwpf.usermodel.XWPFParagraph		// 用于处理文档、表格等中的段落。
org.apache.poi.xwpf.usermodel.XWPFRun			// 使用一组公共属性定义文本区域。
org.apache.poi.xwpf.usermodel.XWPFTable			// 用于处理表格的类。（仅包含文本）
org.apache.poi.xwpf.usermodel.XWPFTableRow		// XWPFTable中的行。
org.apache.poi.xwpf.usermodel.XWPFTableCell		// XWPFTable中的单元格。（表格内容存在单元格中）
org.apache.poi.xwpf.usermodel.XWPFPicture		// 用于处理文档中的图片。
org.apache.poi.xwpf.usermodel.XWPFPictureData	// 原始图片数据。（图片通常保存在/word/media/中）
```

常用接口

```java
org.apache.poi.xwpf.usermodel.Document			// 封装了图片类型属性对应的int值
org.apache.poi.xwpf.usermodel.IBody				// 表示文档不同部分，提供处理的通用方法
org.apache.poi.xwpf.usermodel.IBodyElement		// 
```

## 二、XWPFDocument类

继承关系

```java
java.lang.Object
	org.apache.poi.ooxml.POIXMLDocumentPart
		org.apache.poi.ooxml.POIXMLDocument
			org.apache.poi.xwpf.usermodel.XWPFDocument
```

实现的接口

```java
java.io.Closeable
java.lang.AutoCloseable
org.apache.poi.xwpf.usermodel.Document
org.apache.poi.xwpf.usermodel.IBody
```

### 1、创建`XWPFDocument`文档

创建一个空的`XWPFDocument`对象。

1.1 调用构造器：

```java
public XWPFDocument() {
	super(newPackage());
	onDocumentCreate();
}
```

1.2 示例：

```java
XWPFDocument document = new XWPFDocument();
```

### 2、打开`XWPFDocument`文档

#### 2.1 输入流方式

从输入流读取数据来创建`XWPFDocument`对象。

2.1.1 示例：

```java
String filePath = ""; 			// 文件名
File file = new File(filePath);
InputStream is = new FileInputStream(file);
XWPFDocument document = new XWPFDocument(is);
```

2.1.2 调用构造器：

```java
public XWPFDocument(InputStream is)throws IOException{
	super(PackageHelper.open(is));
	//build a tree of POIXMLDocumentParts, this workbook being the root
	load(XWPFFactory.getInstance());
}
```

#### 2.2 `OPCPackage`方式

从`OPCPackage`读取数据来创建`XWPFDocument`对象。

2.2.1 示例：

```java
String filePath = ""; 			// 文件名
OPCPackage pkg = new OPCPackage(filePath);
XWPFDocument document = new XWPFDocument(pkg);
```

2.2.2 调用构造器：

```java
public XWPFDocument(OPCPackage pkg) throws IOException {
	super(pkg);
    //build a tree of POIXMLDocumentParts, this document being the root
	load(XWPFFactory.getInstance());
}
```

### 3、保存`XWPFDocument`文档

3.1 方法：

```java
public final void write(OutputStream stream);
```

3.2 示例：

```java
String filePath = "Temp.docx";
File file = new File(filePath);
OPCPackage pkg = new OPCPackage(filePath);
XWPFDocument document = new XWPFDocument(pkg);
OutputStream os = new FileOutputStream(file);
document.write(os);
```

### 4、获取`XWPFDcument`文档`BodyElement`

#### 4.1 获取文档所有`BodyElement`

4.1.1 实现`IBody`接口的`getBodyElements()`方法

```java
// 提供一个不可修改的包含文档所有BodyElement对象的列表
@Override
public List<IBodyElement> getBodyElements() {
    return Collections.unmodifiableList(bodyElements);
}
```

示例：

```java
List<IBodyElement> bodyElements = document.getBodyElements();
for (IBodyElement bodyElement : bodyElements) {
	if (bodyElement instanceof XWPFParagraph) {
		XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
        System.out.println("这是一个" + paragraph.getElementType());
    } else if (bodyElement instanceof XWPFTable) {
        XWPFTable table = (XWPFTable) bodyElement;
        System.out.println("这是一个" + table.getElementType());
    } else {
        System.out.println("这是一个" + bodyElement.getElementType());
    }
}
```

4.1.2 `getBodyElementsIterator()`方法

```java
// 提供一个包含文档所有BodyElement对象的迭代器
public Iterator<IBodyElement> getBodyElementsIterator() {
	return bodyElements.iterator();
}
```

示例

```java
Iterator<IBodyElement> bodyElements = document.getBodyElementsIterator();
while (bodyElements.hasNext()) {
    IBodyElement bodyElement = bodyElements.next();
    if (bodyElement instanceof XWPFParagraph) {
        XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
        System.out.println("这是一个" + paragraph.getElementType());
    } else if (bodyElement instanceof XWPFTable) {
        XWPFTable table = (XWPFTable) bodyElement;
        System.out.println("这是一个" + table.getElementType());
    } else {
        System.out.println("这是一个" + bodyElement.getElementType());
    }
}
```

### 5、获取`XWPFDcument`文档的表格

#### 5.1 获取文档所有表格

5.1.1 实现`IBody`接口的`getTables()`方法

```java
@Override
public List<XWPFTable> getTables() {
	return Collections.unmodifiableList(tables);
}
```

示例：

```java
List<XWPFTable> tables = document.getTables();
for (XWPFTable table : tables) {
	int rowCount = table.getNumberOfRows();
	System.out.println(rowCount);
}
```

5.1.2 `getTablesIterator()`方法

```java
public Iterator<XWPFTable> getTablesIterator() {
	return tables.iterator();
}
```

示例：

```java
Iterator<XWPFTable> tables = document.getTablesIterator();
while (tables.hasNext()) {
    XWPFTable table = tables.next();
    List<XWPFTableRow> rows = table.getRows();
    System.out.println(rows.size());
}
```

