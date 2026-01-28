use std::io;

use io::BufWriter;
use io::Write;

use io::BufRead;

use serde_json::Deserializer;
use serde_json::Map;
use serde_json::Number;
use serde_json::Value;

use rust_xlsxwriter::Workbook;
use rust_xlsxwriter::Worksheet;

use rust_xlsxwriter::ColNum;
use rust_xlsxwriter::RowNum;

#[derive(Debug, Clone)]
pub enum XErr {
    UnableToWriteBool(String),
    UnableToWriteDouble(String),
    UnableToWriteString(String),
    UnableToConvertArrayToJson(String),
    UnableToConvertObjectToJson(String),
    InvalidString,
    InvalidNumber(String),
    ColumnOverflow(String),
    RowOverflow(String),
    UnableToParseLine(String),
    InvalidSheetName(String),
    UnableToSaveToBuffer(String),
    UnableToWriteToWriter,
    UnableToFlush,
    EmptyInput,
}

pub struct Sheet<'a> {
    pub ws: &'a mut Worksheet,
}

impl<'a> Sheet<'a> {
    pub fn write_null(&mut self, row: RowNum, col: ColNum) {
        self.ws.clear_cell(row, col);
    }
}

impl<'a> Sheet<'a> {
    pub fn write_bool(&mut self, row: RowNum, col: ColNum, val: bool) -> Result<(), XErr> {
        self.ws
            .write_boolean(row, col, val)
            .map_err(|_| XErr::UnableToWriteBool(format!("rejected value: {val}")))
            .map(|_| ())
    }

    pub fn write_double(&mut self, row: RowNum, col: ColNum, val: f64) -> Result<(), XErr> {
        self.ws
            .write_number(row, col, val)
            .map_err(|_| XErr::UnableToWriteDouble(format!("rejected value: {val}")))
            .map(|_| ())
    }

    pub fn write_number(&mut self, row: RowNum, col: ColNum, val: Number) -> Result<(), XErr> {
        let o: Option<f64> = val.as_f64();
        let f: f64 = o.ok_or_else(|| XErr::InvalidNumber(format!("rejected value: {val}")))?;
        self.write_double(row, col, f)
    }

    pub fn write_string(&mut self, row: RowNum, col: ColNum, val: String) -> Result<(), XErr> {
        self.write_str(row, col, &val)
    }

    pub fn write_str(&mut self, row: RowNum, col: ColNum, val: &str) -> Result<(), XErr> {
        self.ws
            .write_string(row, col, val)
            .map_err(|_| XErr::UnableToWriteString(format!("rejected value: {val}")))
            .map(|_| ())
    }

    pub fn write_array(
        &mut self,
        row: RowNum,
        col: ColNum,
        val: Vec<Value>,
        buf: &mut Vec<u8>,
    ) -> Result<(), XErr> {
        buf.clear();
        serde_json::to_writer(buf.by_ref(), &val)
            .map_err(|e| XErr::UnableToConvertArrayToJson(e.to_string()))?;
        let bs: &[u8] = buf;
        let s: &str = std::str::from_utf8(bs).map_err(|_| XErr::InvalidString)?;
        self.write_str(row, col, s)
    }

    pub fn write_object(
        &mut self,
        row: RowNum,
        col: ColNum,
        val: Map<String, Value>,
        buf: &mut Vec<u8>,
    ) -> Result<(), XErr> {
        buf.clear();
        serde_json::to_writer(buf.by_ref(), &val)
            .map_err(|e| XErr::UnableToConvertObjectToJson(e.to_string()))?;
        let bs: &[u8] = buf;
        let s: &str = std::str::from_utf8(bs).map_err(|_| XErr::InvalidString)?;
        self.write_str(row, col, s)
    }
}

impl<'a> Sheet<'a> {
    pub fn write_value(
        &mut self,
        row: RowNum,
        col: ColNum,
        val: Value,
        buf: &mut Vec<u8>,
    ) -> Result<(), XErr> {
        match val {
            Value::Null => {
                self.write_null(row, col);
                Ok(())
            }
            Value::Bool(val) => self.write_bool(row, col, val),
            Value::Number(val) => self.write_number(row, col, val),
            Value::String(val) => self.write_string(row, col, val),
            Value::Array(val) => self.write_array(row, col, val, buf),
            Value::Object(val) => self.write_object(row, col, val, buf),
        }
    }
}

impl<'a> Sheet<'a> {
    pub fn write_row(
        &mut self,
        row: RowNum,
        val: Map<String, Value>,
        buf: &mut Vec<u8>,
    ) -> Result<(), XErr> {
        for (cno, pair) in val.into_iter().enumerate() {
            let col: u16 = cno
                .try_into()
                .map_err(|_| XErr::ColumnOverflow(format!("rejected column index: {cno}")))?;
            let val: Value = pair.1;
            let colnum: ColNum = col;
            self.write_value(row, colnum, val, buf)?;
        }
        Ok(())
    }

    pub fn write_header<I>(&mut self, headers: I) -> Result<(), XErr>
    where
        I: IntoIterator<Item = String>,
    {
        for (cno, hdr) in headers.into_iter().enumerate() {
            let col: u16 = cno
                .try_into()
                .map_err(|_| XErr::ColumnOverflow(format!("rejected column index: {cno}")))?;
            let key: &String = &hdr;
            let colnum: ColNum = col;
            self.write_str(0, colnum, key)?;
        }
        Ok(())
    }

    pub fn write_rows<I, K>(&mut self, values: I, buf: &mut Vec<u8>, keys: K) -> Result<(), XErr>
    where
        I: Iterator<Item = Result<Map<String, Value>, XErr>>,
        K: IntoIterator<Item = String>,
    {
        self.write_header(keys)?;
        for (rno, rslt) in values.enumerate() {
            let rix: usize = rno + 1; // 1,2,3, ...(0: header)
            let ru: u32 = rix
                .try_into()
                .map_err(|_| XErr::RowOverflow(format!("rejected row index: {rix}")))?;
            let rno: RowNum = ru;

            let val: Map<_, _> = rslt?;
            self.write_row(rno, val, buf)?;
        }
        Ok(())
    }
}

pub fn rdr2jsons<R>(rdr: R) -> impl Iterator<Item = Result<Map<String, Value>, XErr>>
where
    R: BufRead,
{
    Deserializer::from_reader(rdr)
        .into_iter()
        .map(|rslt| rslt.map_err(|e| XErr::UnableToParseLine(e.to_string())))
}

pub struct Book {
    pub wb: Workbook,
}

impl Default for Book {
    fn default() -> Self {
        Self {
            wb: Workbook::new(),
        }
    }
}

impl Book {
    pub fn jsons2sheet<I, K>(
        &mut self,
        jsons: I,
        sheet_name: String,
        buf: &mut Vec<u8>,
        keys: K,
    ) -> Result<(), XErr>
    where
        I: Iterator<Item = Result<Map<String, Value>, XErr>>,
        K: IntoIterator<Item = String>,
    {
        let ws: &mut Worksheet = self.wb.add_worksheet();
        ws.set_name(sheet_name)
            .map_err(|e| XErr::InvalidSheetName(format!("unable to set the sheet name: {e}")))?;
        let mut s = Sheet { ws };
        s.write_rows(jsons, buf, keys)?;
        Ok(())
    }
}

impl Book {
    pub fn save_to_buffer(&mut self) -> Result<Vec<u8>, XErr> {
        self.wb
            .save_to_buffer()
            .map_err(|e| XErr::UnableToSaveToBuffer(e.to_string()))
    }

    pub fn save_to_writer<W>(&mut self, mut w: W) -> Result<(), XErr>
    where
        W: Write,
    {
        let buf: Vec<u8> = self.save_to_buffer()?;
        w.write_all(&buf).map_err(|_| XErr::UnableToWriteToWriter)?;
        w.flush().map_err(|_| XErr::UnableToFlush)
    }
}

pub fn reader2jsons2sheet2writer<R, W>(
    rdr: R,
    sheet_name: String,
    output: W,
    buf: &mut Vec<u8>,
) -> Result<(), XErr>
where
    R: BufRead,
    W: Write,
{
    let jsons = rdr2jsons(rdr); // impl Iterator<Item=Result<Map, XErr>>
    let mut pjsons = jsons.peekable();
    let o1st: Option<&Result<Map<_, _>, _>> = pjsons.peek();
    let rslt1st: &Result<_, _> = o1st.ok_or(XErr::EmptyInput)?;
    let m1st: &Map<_, _> = match rslt1st {
        Ok(m) => Ok(m),
        Err(e) => Err(e.clone()),
    }?;
    let keys: Vec<String> = m1st.keys().cloned().collect();
    let mut bk: Book = Book::default();
    bk.jsons2sheet(pjsons, sheet_name, buf, keys)?;
    bk.save_to_writer(output)
}

pub fn stdin2jsons2sheet2stdout(sheet_name: String, buf: &mut Vec<u8>) -> Result<(), XErr> {
    let o = io::stdout();
    let mut ol = o.lock();
    reader2jsons2sheet2writer(io::stdin().lock(), sheet_name, BufWriter::new(&mut ol), buf)?;
    ol.flush().map_err(|_| XErr::UnableToFlush)
}

pub const SHEET_NAME_DEFAULT: &str = "Sheet1";
pub const BUF_CAP_DEFAULT: usize = 0;

pub fn stdin2jsons2sheet2stdout_default() -> Result<(), XErr> {
    let mut buf: Vec<u8> = Vec::with_capacity(BUF_CAP_DEFAULT);
    stdin2jsons2sheet2stdout(SHEET_NAME_DEFAULT.into(), &mut buf)
}
