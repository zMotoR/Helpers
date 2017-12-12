unit DataSetHelper;

interface

uses Data.DB, System.SysUtils, Xml.XmlIntf, Xml.XmlDoc
{$IFDEF EH_LIB_25}
, MemTableEh
{$ENDIF}
;

type

TDataSetLoopOption = (
  ///	<summary>
  ///	  Запрещает восстановление закладок.
  ///	</summary>
  dsloNoRestoreBookmark,
  ///	<summary>
  ///	  Запрещает отключение элементов управления.
  ///	</summary>
  dsloNoDisableControls,
  ///	<summary>
  ///	  Запрещает отключение фильтров.
  ///	</summary>
  dsloNoDisableFilter,
  ///	<summary>
  ///	  Запрещает обнуление события BeforeScroll.
  ///	</summary>
  dsloNoNullifyBeforeScrollEvent,
  ///	<summary>
  ///	  Запрещает обнуление события AfterScroll.
  ///	</summary>
  dsloNoNullifyAfterScrollEvent,
  ///	<summary>
  ///	  Начинать с текущей записи.
  ///	</summary>
  dsloFromCurrent,
  ///	<summary>
  ///	  Идти в обратном порядке.
  ///	</summary>
  dsloReverseOrder,
  ///	<summary>
  ///	  Только текущая запись.
  ///	</summary>
  dsloOnlyCurrentRecord,
  ///	<summary>
  ///	  Только измененные записи.
  ///	</summary>
  dsloOnlyModifiedRecords
);

TDataSetLoopOptions = set of TDataSetLoopOption;

TDataSetLoopState = class
  private
    broken, restoreBookmark: boolean;
  public
    constructor Create;
    ///	<summary>
    ///	  Останавливает перебор записей объекта TDataSet.
    ///	</summary>
    ///	<param name="restoreBookmark">
    ///	  Нужно ли восстанавливать закладку. По умолчанию true.
    ///	</param>
    procedure Break(restoreBookmark: boolean = true);
    ///	<summary>
    ///	  Возвращает true, если перебор записей остановлен.
    ///	</summary>
    property IsBroken: boolean read broken;
end;

TDataSetHelper = class helper for TDataSet
  public
    ///	<summary>
    ///	  Проверяет, активен ли объект TDataSet и есть ли в нём записи.
    ///	</summary>
    function HasRecords: boolean;

    ///	<summary>
    ///	  Перебирает все записи объекта TDataSet и вызывает для них процедуру proc.
    ///   Во время перебора отключаются элементы управления, фильтр и
    ///   сохраняется закладка.
    ///	</summary>
    ///	<param name="proc">
    ///	  Процедура, которая будет вызвана для каждой записи объекта TDataSet.
    ///	</param>
    procedure ForEach(proc: TProc); overload;

    ///	<summary>
    ///	  Перебирает все записи объекта TDataSet и вызывает для них процедуру proc.
    ///   Во время перебора отключаются элементы управления, фильтр и
    ///   сохраняется закладка.
    ///	</summary>
    ///	<param name="proc">
    ///	  Процедура, которая будет вызвана для каждой записи объекта TDataSet.
    ///	</param>
    procedure ForEach(proc: TProc<TDataSetLoopState>); overload;

    ///	<summary>
    ///	  Перебирает все записи объекта TDataSet и вызывает для них процедуру proc.
    ///	</summary>
    ///	<param name="options">
    ///	  Опции выполнения перебора.
    ///	</param>
    ///	<param name="proc">
    ///	  Процедура, которая будет вызвана для каждой записи объекта TDataSet.
    ///	</param>
    procedure ForEach(options: TDataSetLoopOptions; proc: TProc); overload;

    ///	<summary>
    ///	  Перебирает все записи объекта TDataSet и вызывает для них процедуру proc.
    ///	</summary>
    ///	<param name="options">
    ///	  Опции выполнения перебора.
    ///	</param>
    ///	<param name="proc">
    ///	  Процедура, которая будет вызвана для каждой записи объекта TDataSet.
    ///	</param>
    procedure ForEach(options: TDataSetLoopOptions;
        proc: TProc<TDataSetLoopState>); overload;

    ///	<summary>
    ///	  Записывает данные в XML, хранящиеся в объекте TDataSet.
    ///   Сохраняются все поля.
    ///	</summary>
    ///	<param name="rowPath">
    ///	  Путь к элементам XML, в которые будут записаны отдельные записи из
    ///   объекта TDataSet начиная с корневого элемента XML, например, '/mydata/myrow'.
    ///   Пространства имён не поддерживаются!
    ///   По умолчанию используется путь '/data/row'.
    ///	</param>
    function ToXML(rowPath: string = ''): IXMLDocument; overload;

    ///	<summary>
    ///	  Записывает данные в XML, хранящиеся в объекте TDataSet.
    ///	</summary>
    ///	<param name="rowPath">
    ///	  Путь к элементам XML, в которые будут записаны отдельные записи из
    ///   объекта TDataSet начиная с корневого элемента XML, например, '/mydata/myrow'.
    ///   Пространства имён не поддерживаются!
    ///   По умолчанию используется путь '/data/row'.
    ///	</param>
    ///	<param name="fields">
    ///	  Имена полей для сохранения в XML.
    ///   Если массив пустой, то идёт сохранение всех полей.
    ///	</param>
    function ToXML(rowPath: string; fields: array of string)
        : IXMLDocument; overload;

    ///	<summary>
    ///	  Записывает данные в XML, хранящиеся в объекте TDataSet.
    ///   Для записей используется путь '/data/row'.
    ///   Сохраняются все поля.
    ///	</summary>
    ///	<param name="fields">
    ///	  Имена полей для сохранения в XML.
    ///   Если массив пустой, то идёт сохранение всех полей.
    ///	</param>
    function ToXML(fields: array of string): IXMLDocument; overload;

    ///	<summary>
    ///	  Записывает данные в XML, хранящиеся в объекте TDataSet.
    ///   Если XML не указан, то создаёт новый экземпляр IXMLDocument.
    ///   Сохраняются все поля.
    ///	</summary>
    ///	<param name="rowPath">
    ///	  Путь к элементам XML, в которые будут записаны отдельные записи из
    ///   объекта TDataSet начиная с корневого элемента XML, например, '/mydata/myrow'.
    ///   Пространства имён не поддерживаются!
    ///   По умолчанию используется путь '/data/row'.
    ///	</param>
    ///	<param name="doc">
    ///	  Документ, внутрь которого нужно записать данные.
    ///   Если указан nil, то будет создан новый документ.
    ///	</param>
    ///	<returns>
    ///	  Возвращает XML-документ с записанными в него данными.
    ///	</returns>
    function ToXML(rowPath: string; doc: IXMLDocument): IXMLDocument; overload;

    ///	<summary>
    ///	  Записывает данные в XML, хранящиеся в объекте TDataSet.
    ///   Если XML не указан, то создаёт новый экземпляр IXMLDocument.
    ///   Сохраняются все поля.
    ///	</summary>
    ///	<param name="doc">
    ///	  Документ, внутрь которого нужно записать данные.
    ///   Если указан nil, то будет создан новый документ.
    ///	</param>
    ///	<returns>
    ///	  Возвращает XML-документ с записанными в него данными.
    ///	</returns>
    function ToXML(doc: IXMLDocument): IXMLDocument; overload;

    ///	<summary>
    ///	  Записывает данные в XML, хранящиеся в объекте TDataSet.
    ///   Если XML не указан, то создаёт новый экземпляр IXMLDocument.
    ///	</summary>
    ///	<param name="rowPath">
    ///	  Путь к элементам XML, в которые будут записаны отдельные записи из
    ///   объекта TDataSet начиная с корневого элемента XML, например, '/mydata/myrow'.
    ///   Пространства имён не поддерживаются!
    ///   По умолчанию используется путь '/data/row'.
    ///	</param>
    ///	<param name="doc">
    ///	  Документ, внутрь которого нужно записать данные.
    ///   Если указан nil, то будет создан новый документ.
    ///	</param>
    ///	<param name="fields">
    ///   Имена полей для сохранения в XML.
    ///   Если массив пустой, то идёт сохранение всех полей.
    ///	</param>
    ///	<returns>
    ///	  Возвращает XML-документ с записанными в него данными.
    ///	</returns>
    function ToXML(rowPath: string; doc: IXMLDocument;
        fields: array of string) : IXMLDocument; overload;

    ///	<summary>
    ///	  Записывает данные в XML, хранящиеся в объекте TDataSet.
    ///   Если XML не указан, то создаёт новый экземпляр IXMLDocument.
    ///	</summary>
    ///	<param name="doc">
    ///	  Документ, внутрь которого нужно записать данные.
    ///   Если указан nil, то будет создан новый документ.
    ///	</param>
    ///	<param name="fields">
    ///   Имена полей для сохранения в XML.
    ///   Если массив пустой, то идёт сохранение всех полей.
    ///	</param>
    ///	<returns>
    ///	  Возвращает XML-документ с записанными в него данными.
    ///	</returns>
    function ToXML(doc: IXMLDocument; fields: array of string)
        : IXMLDocument; overload;

    ///	<summary>
    ///	  Записывает данные в XML, хранящиеся в объекте TDataSet.
    ///   Если параметр doc = nil, то создаёт новый экземпляр IXMLDocument.
    ///	</summary>
    ///	<param name="options">
    ///	  Опции выполнения перебора.
    ///	</param>
    ///	<param name="rowPath">
    ///	  Путь к элементам XML, в которые будут записаны отдельные записи из
    ///   объекта TDataSet начиная с корневого элемента XML, например, '/mydata/myrow'.
    ///   Пространства имён не поддерживаются!
    ///   По умолчанию используется путь '/data/row'.
    ///	</param>
    ///	<param name="doc">
    ///	  Документ, внутрь которого нужно записать данные.
    ///   Если указан nil, то будет создан новый документ.
    ///	</param>
    ///	<param name="fieldNames">
    ///   Имена полей для сохранения в XML.
    ///   Если массив пустой, то идёт сохранение всех полей.
    ///	</param>
    function ToXML(const options: TDataSetLoopOptions; rowPath: string;
        doc: IXMLDocument; fieldNames: array of string): IXMLDocument; overload;
end;

TFieldHelper = class helper for TField
  public
    ///	<summary>
    ///	  Считает сумму значений поля во всех записях.
    ///   Значения null пропускаются.
    ///	</summary>
    ///	<returns>
    ///	  Возвращает сумму значений.
    ///   Если записей нет или все поля null, то возвращает 0.
    ///	</returns>
    function Sum: variant; overload;
    ///	<summary>
    ///	  Считает сумму значений поля с учётом опций перебора.
    ///   Значения null пропускаются.
    ///	</summary>
    ///	<param name="options">
    ///	  Опции выполнения перебора.
    ///	</param>
    ///	<returns>
    ///	  Возвращает сумму значений.
    ///   Если записей нет или все поля null, то возвращает 0.
    ///	</returns>
    function Sum(options: TDataSetLoopOptions): variant; overload;
    ///	<summary>
    ///	  Вычисляет среднее арифметическое значений поля во всех записях.
    ///   Значения null пропускаются.
    ///	</summary>
    ///	<returns>
    ///	  Возвращает среднее арифметическое значений.
    ///   Если записей нет или все поля null, то возвращает 0.
    ///	</returns>
    function Avg: variant; overload;
    ///	<summary>
    ///	  Вычисляет среднее арифметическое значений поля с учётом опций перебора.
    ///   Значения null пропускаются.
    ///	</summary>
    ///	<param name="options">
    ///	  Опции выполнения перебора.
    ///	</param>
    ///	<returns>
    ///	  Возвращает среднее арифметическое значений.
    ///   Если записей нет или все поля null, то возвращает 0.
    ///	</returns>
    function Avg(options: TDataSetLoopOptions): variant; overload;
    ///	<summary>
    ///	  Находит минимальное значение поля во всех записях.
    ///   Значения null пропускаются.
    ///	</summary>
    ///	<returns>
    ///	  Возвращает минимальное значение поля.
    ///   Если записей нет или все поля null, то возвращает 0.
    ///	</returns>
    function Min: variant; overload;
    ///	<summary>
    ///	  Находит минимальное значение поля с учётом опций перебора.
    ///   Значения null пропускаются.
    ///	</summary>
    ///	<param name="options">
    ///	  Опции выполнения перебора.
    ///	</param>
    ///	<returns>
    ///	  Возвращает минимальное значение поля.
    ///   Если записей нет или все поля null, то возвращает 0.
    ///	</returns>
    function Min(options: TDataSetLoopOptions): variant; overload;
    ///	<summary>
    ///	  Находит максимальное значение поля во всех записях.
    ///   Значения null пропускаются.
    ///	</summary>
    ///	<returns>
    ///	  Возвращает максимальное значение поля.
    ///   Если записей нет или все поля null, то возвращает null.
    ///	</returns>
    function Max: variant; overload;
    ///	<summary>
    ///	  Находит максимальное значение поля с учётом опций перебора.
    ///   Значения null пропускаются.
    ///	</summary>
    ///	<param name="options">
    ///	  Опции выполнения перебора.
    ///	</param>
    ///	<returns>
    ///	  Возвращает максимальное значение поля.
    ///   Если записей нет или все поля null, то возвращает null.
    ///	</returns>
    function Max(options: TDataSetLoopOptions): variant; overload;
end;

///	<summary>
///	  Функция возвращает значение ATrue, если AValue равно True и AFalse, если
///	  AValue равно False.
///	</summary>
///	<param name="AValue">
///	  Флаг, определяющий, какое значение нужно вернуть.
///	</param>
///	<param name="ATrue">
///	  Значение, которое нужно вернуть, если аргумент AValue равен True.
///	</param>
///	<param name="AFalse">
///	  Значение, которое нужно вернуть, если аргумент AValue равен False.
///	</param>
function IfThenVar(AValue: Boolean; const ATrue: variant; const AFalse: variant)
    : variant;

///	<summary>
///	  Функция конвертирует значение с типом Variant в строку с учётом локали.
///	</summary>
///	<param name="v">
///	  Конвертируемое значение.
///	</param>
///	<param name="AFormatSettings">
///	  Локальные настройки.
///	</param>
function VarToStrLoc(const v: variant; const formatSettings: TFormatSettings)
    : string;

///	<summary>
///	  Конвертирует значение с типом Variant в строку для использования в XML.
///	</summary>
///	<param name="v">
///	  Конвертируемое значение.
///	</param>
///	<param name="dataType">
///	  Тип данных. Используется только для даты и времени.
///	</param>
function VarToXmlStr(const v: variant; dataType: TFieldType = ftDateTime)
    : string;

implementation

uses System.Variants, Soap.XSBuiltIns, Winapi.Windows;

var
  xmlFormatSettings: TFormatSettings;

constructor TDataSetLoopState.Create;
begin
  broken := false;
  restoreBookmark := true;
end;

procedure TDataSetLoopState.Break(restoreBookmark: boolean = true);
begin
  broken := true;
  self.restoreBookmark := restoreBookmark;
end;

procedure TDataSetHelper.ForEach(proc: TProc);
begin
  ForEach([], proc);
end;

procedure TDataSetHelper.ForEach(proc: TProc<TDataSetLoopState>);
begin
  ForEach([], proc);
end;

procedure TDataSetHelper.ForEach(options: TDataSetLoopOptions; proc: TProc);
begin
  ForEach(options, procedure (state: TDataSetLoopState)
    begin
      proc;
    end);
end;

procedure TDataSetHelper.ForEach(options: TDataSetLoopOptions;
    proc: TProc<TDataSetLoopState>);
var
  state: TDataSetLoopState;
  cd, fltr: boolean;
  bmrk: TBookmark;
  beforeScrollEvent, afterScrollEvent: TDataSetNotifyEvent;
begin
  if not assigned(self) then
    raise EAccessViolation.Create('Набор данных не существует.');
  if Active then
  begin
    beforeScrollEvent := BeforeScroll;
    afterScrollEvent := AfterScroll;
    if not (dsloNoNullifyBeforeScrollEvent in options) then
      BeforeScroll := nil;
    try
      if not (dsloNoNullifyAfterScrollEvent in options) then
        AfterScroll := nil;
      try
        state := TDataSetLoopState.Create;
        try
          cd := ControlsDisabled;
          if (not cd) and (not (dsloNoDisableControls in options)) then
            DisableControls;
          try
            bmrk := Bookmark;
            try
              fltr := Filtered and (not (dsloNoDisableFilter in options));
              if fltr then
                Filtered := false;
              try
                if RecordCount > 0 then
                begin
                  if not (dsloFromCurrent in options) then
                    if dsloReverseOrder in options then
                      Last
                    else
                      First;
                  while ((dsloReverseOrder in options) and (not Bof))
                    or ((not (dsloReverseOrder in options)) and (not Eof)) do
                  begin
                    if (not (dsloOnlyModifiedRecords in options))
                        or (UpdateStatus <> usUnmodified) then
                      proc(state);
                    if state.broken or (dsloOnlyCurrentRecord in options) then
                      break;
                    if dsloReverseOrder in options then
                      Prior
                    else
                      Next;
                  end;
                end;
              finally
                if fltr then
                  Filtered := true;
              end;
            finally
              if (not (dsloNoRestoreBookmark in options))
                  and state.restoreBookmark then
              begin
                {$IFDEF EH_LIB_25}
                if self is TMemTableEh then
                begin
                  if TMemTableEh(self).BookmarkInVisibleView(bmrk) then
                    Bookmark := bmrk;
                end
                else
                {$ENDIF}
                if BookmarkValid(bmrk) then
                  Bookmark := bmrk;
              end;
            end;
          finally
            if (not cd) and (not (dsloNoDisableControls in options)) then
              EnableControls;
          end;
        finally
          state.Free;
        end;
      finally
        if not (dsloNoNullifyAfterScrollEvent in options) then
          AfterScroll := afterScrollEvent;
      end;
    finally
      if not (dsloNoNullifyBeforeScrollEvent in options) then
        BeforeScroll := beforeScrollEvent;
    end;
  end;
end;

function TDataSetHelper.HasRecords: boolean;
begin
  Result := Active and (RecordCount > 0);
end;

function TDataSetHelper.ToXML(rowPath: string = ''): IXMLDocument;
begin
  Result := ToXML([], rowPath, nil, []);
end;

function TDataSetHelper.ToXML(rowPath: string; fields: array of string)
    : IXMLDocument;
begin
  Result := ToXML([], rowPath, nil, fields);
end;

function TDataSetHelper.ToXML(fields: array of string): IXMLDocument;
begin
  Result := ToXML([], '', nil, fields);
end;

function TDataSetHelper.ToXML(rowPath: string; doc: IXMLDocument)
    : IXMLDocument;
begin
  Result := ToXML([], rowPath, doc, []);
end;

function TDataSetHelper.ToXML(doc: IXMLDocument): IXMLDocument;
begin
  Result := ToXML([], '', doc, []);
end;

function TDataSetHelper.ToXML(rowPath: string; doc: IXMLDocument;
    fields: array of string) : IXMLDocument;
begin
  Result := ToXML([], rowPath, doc, fields);
end;

function TDataSetHelper.ToXML(doc: IXMLDocument; fields: array of string)
    : IXMLDocument;
begin
  Result := ToXML([], '', doc, fields);
end;

function TDataSetHelper.ToXML(const options: TDataSetLoopOptions; rowPath: string;
    doc: IXMLDocument; fieldNames: array of string): IXMLDocument;
var
  path: TArray<string>;
  node, parentNode: IXMLNode;
  i: integer;
  fields: array of TField;
  rowTagName: string;
begin
  if not assigned(self) then
    raise EAccessViolation.Create('Набор данных не существует.');
  if rowPath.IsEmpty then
    if assigned(doc) and assigned(doc.DocumentElement) then
      rowPath := '/' + doc.DocumentElement.NodeName + '/row'
    else
      rowPath := '/data/row';
  path := rowPath.Split(['/'], TStringSplitOptions.ExcludeEmpty);
  if Length(path) < 2 then
    raise Exception.Create('Путь должен состоять как минимум из двух элементов XML, например, ''/data/row''.');
  if not assigned(doc) then
    doc := TXMLDocument.Create(nil);
  if not doc.Active then
    doc.Active := true;
  if not assigned(doc.DocumentElement) then
    doc.DocumentElement := doc.CreateElement(path[0], '')
  else if path[0] <> doc.DocumentElement.NodeName then
    raise Exception.Create('Имя корневого элемента указанное в пути ''' + path[0]
        + ''' и имя корневого элемента в XML-документе '''
        + doc.DocumentElement.NodeName + ''' не соответствуют.');
  parentNode := doc.DocumentElement;
  for i := 1 to Length(path) - 2 do
    parentNode := parentNode.ChildNodes.Nodes[path[i]];
  rowTagName := path[Length(path) - 1];
  if Length(fieldNames) = 0 then
  begin
    SetLength(fields, FieldCount);
    for i := 0 to FieldCount - 1 do
      fields[i] := self.Fields[i];
  end
  else
  begin
    SetLength(fields, Length(fieldNames));
    for i := 0 to Length(fieldNames) - 1 do
      fields[i] := self.FieldByName(fieldNames[i]);
  end;
  ForEach(options, procedure
    var
      field: TField;
    begin
      node := parentNode.AddChild(rowTagName);
      for field in fields do
        if not field.IsNull then
          node.SetAttribute(field.FieldName,
              VarToXmlStr(field.Value, field.DataType));
    end
  );
  Result := doc;
end;

function TFieldHelper.Sum: variant;
begin
  Result := Sum([]);
end;

function TFieldHelper.Sum(options: TDataSetLoopOptions): variant;
var
  s: variant;
begin
  if not assigned(self) then
    raise EAccessViolation.Create('Поле не существует.');
  s := null;
  DataSet.ForEach(options, procedure
    begin
      if not IsNull then
        if VarIsNull(s) then
          s := Value
        else
          s := s + Value;
    end
  );
  if VarIsNull(s) then
    s := 0;
  Result := s;
end;

function TFieldHelper.Avg: variant;
begin
  Result := Avg([]);
end;

function TFieldHelper.Avg(options: TDataSetLoopOptions): variant;
var
  s, c: variant;
begin
  if not assigned(self) then
    raise EAccessViolation.Create('Поле не существует.');
  Result := 0;
  c := 0;
  s := null;
  DataSet.ForEach(options, procedure
    begin
      if not IsNull then
      begin
        if VarIsNull(s) then
          s := Value
        else
          s := s + Value;
        Inc(c);
      end;
    end
  );
  if c > 0 then
    Result := s / c;
end;

function TFieldHelper.Min: variant;
begin
  Result := Min([]);
end;

function TFieldHelper.Min(options: TDataSetLoopOptions): variant;
var
  m: variant;
begin
  if not assigned(self) then
    raise EAccessViolation.Create('Поле не существует.');
  m := null;
  DataSet.ForEach(options, procedure
    begin
      if not IsNull then
        if VarIsNull(m) then
          m := Value
        else
          m := IfThenVar(m < Value, m, Value);
    end
  );
  Result := m;
end;

function TFieldHelper.Max: variant;
begin
  Result := Max([]);
end;

function TFieldHelper.Max(options: TDataSetLoopOptions): variant;
var
  m: variant;
begin
  if not assigned(self) then
    raise EAccessViolation.Create('Поле не существует.');
  m := null;
  DataSet.ForEach(options, procedure
    begin
      if not IsNull then
        if VarIsNull(m) then
          m := Value
        else
          m := IfThenVar(m > Value, m, Value);
    end
  );
  Result := m;
end;

function IfThenVar(AValue: Boolean; const ATrue: variant; const AFalse: variant): variant;
begin
  if AValue then
    result := ATrue
  else
    result := AFalse;
end;

function VarToStrLoc(const v: variant; const formatSettings: TFormatSettings): string;
begin
  case VarType(v) of
    varSingle:
      Result := FloatToStr(v, formatSettings);
    varDouble:
      Result := FloatToStr(v, formatSettings);
    varCurrency:
      Result := CurrToStr(v, formatSettings);
    varDate:
      Result := DateTimeToStr(v, formatSettings);
  else
    Result := System.Variants.VarToStr(v);
  end;
end;

function VarToXmlStr(const v: variant; dataType: TFieldType): string;
begin
  case VarType(v) of
    varDate:
      case dataType of
        ftDate:
          Result := FormatDateTime(xmlFormatSettings.ShortDateFormat, v, xmlFormatSettings);//ISO8601
        ftTime:
          Result := FormatDateTime(xmlFormatSettings.ShortTimeFormat, v, xmlFormatSettings);//ISO8601
      else
        //ISO8601
        Result := DateTimeToXMLTime(v, false);//без пояса
      end;
  else
    Result := VarToStrLoc(v, xmlFormatSettings);
  end;
end;

initialization
  xmlFormatSettings := TFormatSettings.Create;
  xmlFormatSettings.DecimalSeparator := '.';
  xmlFormatSettings.DateSeparator := '-';
  xmlFormatSettings.ShortDateFormat := 'yyyy-mm-dd';//ISO8601
  xmlFormatSettings.LongDateFormat := 'yyyy-mm-dd';//ISO8601
  xmlFormatSettings.ShortTimeFormat := 'hh":"nn":"ss.zzz';//ISO8601
  xmlFormatSettings.LongTimeFormat := 'hh":"nn":"ss.zzz';//ISO8601

end.
