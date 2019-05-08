using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Configuration;
using System.Web;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;

namespace Excel操作
{
    public class ClassXML
    {
        public string txtname;

        public void ready(StreamWriter sw,int MNum,int ID)//打开一个Excel文件
        {

            string dpm = "\"";//  double quotation marks

            sw.Write("\r\n" + @"<SW.Blocks.CompileUnit ID=" + dpm + (ID+1) + dpm + @" CompositionName=" + dpm + "CompileUnits" + dpm + @">"); //程序段
            sw.Write("\r\n" + @" <AttributeList>");
            sw.Write("\r\n" + @"<NetworkSource><FlgNet xmlns=" + dpm + "http://www.siemens.com/automation/Openness/SW/NetworkSource/FlgNet/v1" + dpm + ">");
            sw.Write("\r\n" + @"<Parts>");   //当前程序段符号定义 每一个程序段内的UID不能重复 
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "GlobalVariable" + dpm + " UId=" + dpm + "21" + dpm + @">"); //定义符号的uid  全局变量 Input.M[101].BQ1
            sw.Write("\r\n" + @"<Symbol>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "Input" + dpm + @"/>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "M" + dpm + @">");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "101" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"</Component>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "BQ1" + dpm + @"/>");
            sw.Write("\r\n" + @"<Address Area=" + dpm + "DB" + dpm + " Type=" + dpm + "Bool" + dpm + " BlockNumber=" + dpm + "11" + dpm + " BitOffset=" + dpm + "41" + dpm + " Informative=" + dpm + "true" + dpm + @"/>");
            sw.Write("\r\n" + @"</Symbol>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "TypedConstant" + dpm + " UId=" + dpm + "22" + dpm + @">"); //定义符号的uid 时间变量ton
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantValue>T#2s</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Time</StringAttribute>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "FormatFlags" + dpm + " Informative=" + dpm + "true" + dpm + @">TypeQualifier</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "GlobalVariable" + dpm + " UId=" + dpm + "23" + dpm + @">"); //定义符号的uid  全局变量 Input.M[101].BQ1
            sw.Write("\r\n" + @"<Symbol>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "STA" + dpm + @"/>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "M" + dpm + @">");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "101" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"</Component>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "Fault" + dpm + @"/>");
            sw.Write("\r\n" + @"<Address Area=" + dpm + "DB" + dpm + " Type=" + dpm + "Bool" + dpm + " BlockNumber=" + dpm + "4" + dpm + " BitOffset=" + dpm + "24" + dpm + " Informative=" + dpm + "true" + dpm + @"/>");
            sw.Write("\r\n" + @"</Symbol>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "GlobalVariable" + dpm + " UId=" + dpm + "24" + dpm + @">"); //定义符号的uid  全局变量 Input.M[101].BQ1
            sw.Write("\r\n" + @"<Symbol>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "INFO" + dpm + @"/>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "M" + dpm + @">");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "101" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"</Component>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "Work_ID" + dpm + @"/>");
            sw.Write("\r\n" + @"<Address Area=" + dpm + "DB" + dpm + " Type=" + dpm + "DInt" + dpm + " BlockNumber=" + dpm + "8" + dpm + " BitOffset=" + dpm + "160" + dpm + " Informative=" + dpm + "true" + dpm + @"/>");
            sw.Write("\r\n" + @"</Symbol>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + " UId=" + dpm + "25" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "0" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "GlobalVariable" + dpm + " UId=" + dpm + "26" + dpm + @">"); //定义符号的uid  全局变量 Input.M[101].BQ1
            sw.Write("\r\n" + @"<Symbol>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "Control" + dpm + @"/>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "M" + dpm + @">");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "101" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"</Component>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "Ready" + dpm + @"/>");
            sw.Write("\r\n" + @"<Address Area=" + dpm + "DB" + dpm + " Type=" + dpm + "Bool" + dpm + " BlockNumber=" + dpm + "13" + dpm + " BitOffset=" + dpm + "17" + dpm + " Informative=" + dpm + "true" + dpm + @"/>");
            sw.Write("\r\n" + @"</Symbol>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "GlobalVariable" + dpm + " UId=" + dpm + "27" + dpm + @">"); //定义符号的uid  全局变量 Input.M[101].BQ1
            sw.Write("\r\n" + @"<Symbol>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "STA" + dpm + @"/>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "M" + dpm + @">");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "101" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"</Component>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "Fault" + dpm + @"/>");
            sw.Write("\r\n" + @"<Address Area=" + dpm + "DB" + dpm + " Type=" + dpm + "Bool" + dpm + " BlockNumber=" + dpm + "4" + dpm + " BitOffset=" + dpm + "24" + dpm + " Informative=" + dpm + "true" + dpm + @"/>");
            sw.Write("\r\n" + @"</Symbol>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "GlobalVariable" + dpm + " UId=" + dpm + "28" + dpm + @">"); //定义符号的uid  全局变量 Input.M[101].BQ1
            sw.Write("\r\n" + @"<Symbol>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "INFO" + dpm + @"/>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "M" + dpm + @">");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "101" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"</Component>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "Work_ID" + dpm + @"/>");
            sw.Write("\r\n" + @"<Address Area=" + dpm + "DB" + dpm + " Type=" + dpm + "DInt" + dpm + " BlockNumber=" + dpm + "8" + dpm + " BitOffset=" + dpm + "160" + dpm + " Informative=" + dpm + "true" + dpm + @"/>");
            sw.Write("\r\n" + @"</Symbol>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + " UId=" + dpm + "29" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "0" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "GlobalVariable" + dpm + " UId=" + dpm + "30" + dpm + @">"); //定义符号的uid  全局变量 Input.M[101].BQ1
            sw.Write("\r\n" + @"<Symbol>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "Control" + dpm + @"/>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "M" + dpm + @">");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "101" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"</Component>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "Ask_F" + dpm + @"/>");
            sw.Write("\r\n" + @"<Address Area=" + dpm + "DB" + dpm + " Type=" + dpm + "Bool" + dpm + " BlockNumber=" + dpm + "13" + dpm + " BitOffset=" + dpm + "18" + dpm + " Informative=" + dpm + "true" + dpm + @"/>");
            sw.Write("\r\n" + @"</Symbol>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"<Part Name=" + dpm + "Contact" + dpm + " UId=" + dpm + "31" + dpm + @">"); //闭点线圈  //程线圈程序定义
            sw.Write("\r\n" + @"<Negated Name=" + dpm + "operand" + dpm + @"/>");
            sw.Write("\r\n" + @"</Part>");
            sw.Write("\r\n" + @"<Part Name=" + dpm + "TOF" + dpm + " Version=" + dpm + "1.0" + dpm + " UId=" + dpm + "32" + dpm + @">"); //时间
            sw.Write("\r\n" + @"<Instance UId=" + dpm + "33" + dpm + " Scope=" + dpm + "GlobalVariable" + dpm + @">"); //
            sw.Write("\r\n" + @"<Component Name=" + dpm + "DB2T" + dpm + @"/>");
            sw.Write("\r\n" + @"<Component Name=" + dpm + "T" + dpm + @">");
            sw.Write("\r\n" + @"<Access Scope=" + dpm + "LiteralConstant" + dpm + @">");
            sw.Write("\r\n" + @"<Constant>");
            sw.Write("\r\n" + @"<ConstantType>DInt</ConstantType>");
            sw.Write("\r\n" + @"<ConstantValue>" + "101" + "</ConstantValue>");
            sw.Write("\r\n" + @"<StringAttribute Name=" + dpm + "Format" + dpm + " Informative=" + dpm + "true" + dpm + @">Dec_signed</StringAttribute>");
            sw.Write("\r\n" + @"</Constant>");
            sw.Write("\r\n" + @"</Access>");
            sw.Write("\r\n" + @"</Component>");
            sw.Write("\r\n" + @"<Address Area=" + dpm + "None" + dpm + " BlockNumber=" + dpm + "2" + dpm + " BitOffset=" + dpm + "12960" + dpm + " Informative=" + dpm + "true" + dpm + @"/>");
            sw.Write("\r\n" + @"</Instance>");
            sw.Write("\r\n" + @" <TemplateValue Name=" + dpm + "time_type" + dpm + " Type=" + dpm + "Type" + dpm + @">Time</TemplateValue>");
            sw.Write("\r\n" + @"</Part>");
            sw.Write("\r\n" + @"<Part Name=" + dpm + "Contact" + dpm + " UId=" + dpm + "34" + dpm + @">"); //闭点线圈
            sw.Write("\r\n" + @"<Negated Name=" + dpm + "operand" + dpm + @"/>");
            sw.Write("\r\n" + @"</Part>");
            sw.Write("\r\n" + @"<Part Name=" + dpm + "Eq" + dpm + " UId=" + dpm + "35" + dpm + @">"); //等于
            sw.Write("\r\n" + @" <TemplateValue Name=" + dpm + "SrcType" + dpm + " Type=" + dpm + "Type" + dpm + @">DInt</TemplateValue>");
            sw.Write("\r\n" + @"</Part>");
            sw.Write("\r\n" + @"<Part Name=" + dpm + "Coil" + dpm + " UId=" + dpm + "36" + dpm + @"/>"); //输出线圈
            sw.Write("\r\n" + @"<Part Name=" + dpm + "Contact" + dpm + " UId=" + dpm + "37" + dpm + @">"); //闭点线圈
            sw.Write("\r\n" + @"<Negated Name=" + dpm + "operand" + dpm + @"/>");
            sw.Write("\r\n" + @"</Part>");
            sw.Write("\r\n" + @"<Part Name=" + dpm + "Ne" + dpm + " UId=" + dpm + "38" + dpm + @">"); //不等于
            sw.Write("\r\n" + @" <TemplateValue Name=" + dpm + "SrcType" + dpm + " Type=" + dpm + "Type" + dpm + @">DInt</TemplateValue>");
            sw.Write("\r\n" + @"</Part>");
            sw.Write("\r\n" + @"<Part Name=" + dpm + "Coil" + dpm + " UId=" + dpm + "39" + dpm + @"/>"); //输出线圈
            sw.Write("\r\n" + @"</Parts>");
            sw.Write("\r\n" + @"<Wires>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "41" + dpm + @">");
            sw.Write("\r\n" + @"<Powerrail />");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "31" + dpm + @" Name=" + dpm + "in" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "37" + dpm + @" Name=" + dpm + "in" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "42" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "21" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "31" + dpm + @" Name=" + dpm + "operand" + dpm + @" />");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "43" + dpm + @">");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "31" + dpm + @" Name=" + dpm + "out" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "32" + dpm + @" Name=" + dpm + "IN" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "44" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "22" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "32" + dpm + @" Name=" + dpm + "PT" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "45" + dpm + @">");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "32" + dpm + @" Name=" + dpm + "Q" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "34" + dpm + @" Name=" + dpm + "in" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "46" + dpm + @">");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "32" + dpm + @" Name=" + dpm + "ET" + dpm + @"/>");
            sw.Write("\r\n" + @"<OpenCon UId=" + dpm + "40" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "47" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "23" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "34" + dpm + @" Name=" + dpm + "operand" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "48" + dpm + @">");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "34" + dpm + @" Name=" + dpm + "out" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "35" + dpm + @" Name=" + dpm + "pre" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "49" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "24" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "35" + dpm + @" Name=" + dpm + "in1" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "50" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "25" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "35" + dpm + @" Name=" + dpm + "in2" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "51" + dpm + @">");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "35" + dpm + @" Name=" + dpm + "out" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "36" + dpm + @" Name=" + dpm + "in" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "52" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "26" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "36" + dpm + @" Name=" + dpm + "operand" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "53" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "27" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "37" + dpm + @" Name=" + dpm + "operand" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "54" + dpm + @">");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "37" + dpm + @" Name=" + dpm + "out" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "38" + dpm + @" Name=" + dpm + "pre" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "55" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "28" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "38" + dpm + @" Name=" + dpm + "in1" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "56" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "29" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "38" + dpm + @" Name=" + dpm + "in2" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "57" + dpm + @">");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "38" + dpm + @" Name=" + dpm + "out" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "39" + dpm + @" Name=" + dpm + "in" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"<Wire UId=" + dpm + "58" + dpm + @">");
            sw.Write("\r\n" + @"<IdentCon UId=" + dpm + "30" + dpm + @"/>");
            sw.Write("\r\n" + @"<NameCon UId=" + dpm + "39" + dpm + @" Name=" + dpm + "operand" + dpm + @"/>");
            sw.Write("\r\n" + @"</Wire>");
            sw.Write("\r\n" + @"</Wires>");
            sw.Write("\r\n" + @"</FlgNet></NetworkSource>");
            sw.Write("\r\n" + @"<ProgrammingLanguage>LAD</ProgrammingLanguage>");
            sw.Write("\r\n" + @"</AttributeList>");
            sw.Write("\r\n" + @"<ObjectList>");
            sw.Write("\r\n" + @"<MultilingualText ID=" + dpm + (ID+2) + dpm + @" CompositionName=" + dpm + "Comment" + dpm + @">");
            sw.Write("\r\n" + @"<ObjectList>");
            sw.Write("\r\n" + @" <MultilingualTextItem ID=" + dpm +( ID+2)+ dpm + @" CompositionName=" + dpm + "Items" + dpm + @">");
            sw.Write("\r\n" + @"<AttributeList>");
            sw.Write("\r\n" + @"<Culture>zh-CN</Culture>");
            sw.Write("\r\n" + @"<Text />");
            sw.Write("\r\n" + @"</AttributeList>");
            sw.Write("\r\n" + @"</MultilingualTextItem>");
            sw.Write("\r\n" + @" </ObjectList>");
            sw.Write("\r\n" + @" </MultilingualText>");
            sw.Write("\r\n" + @" <MultilingualText ID=" + dpm + (ID+4) + dpm + @" CompositionName=" + dpm + "Title" + dpm + @">");
            sw.Write("\r\n" + @"<ObjectList>");
            sw.Write("\r\n" + @" <MultilingualTextItem  ID=" + dpm + (ID + 5) + dpm + @" CompositionName=" + dpm + "Items" + dpm + @">");
            sw.Write("\r\n" + @"<AttributeList>");
            sw.Write("\r\n" + @" <Culture>zh-CN</Culture>");
            sw.Write("\r\n" + @"<Text>" + "Ready101" + "</Text>");//程序注释内容
            sw.Write("\r\n" + @"</AttributeList>");
            sw.Write("\r\n" + @"</MultilingualTextItem>");
            sw.Write("\r\n" + @"</ObjectList>");
            sw.Write("\r\n" + @"</MultilingualText>");
            sw.Write("\r\n" + @"</ObjectList>");
            sw.Write("\r\n" + @"</SW.Blocks.CompileUnit>");













        }
        public void run1(string FileName)//打开一个Excel文件
        {

        }

        public void run2(string FileName)//打开一个Excel文件
        {

        }

    }

    }

