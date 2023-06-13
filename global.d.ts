interface Options {
  keyCode: string;
  valueCode: string;
}

interface XlsxToJsonOption extends Options {
  filePath: string;
}

interface JsonToXlsxOption extends Options {
  dirPath: string;
}
