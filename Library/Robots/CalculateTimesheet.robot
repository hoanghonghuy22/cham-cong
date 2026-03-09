<?xml version="1.0" encoding="UTF-8" ?>
<object class="Robot" serializationversion="2">
  <prologue>
    <saved-by-versions>
      <version>10.4.0.0</version>
    </saved-by-versions>
    <referenced-types>
      <type name="CheckinRecord"/>
    </referenced-types>
    <referenced-snippets/>
    <triggers/>
    <typed-variables>
      <typed-variable name="checkinRecord" type-name="CheckinRecord"/>
    </typed-variables>
    <global-variables>
      <variable name="checkinRecord"/>
    </global-variables>
    <parameters/>
    <return-variables/>
    <store-in-database-variables/>
    <device-mappings/>
    <browser-engine>WEBKIT</browser-engine>
  </prologue>

  <property name="variables" class="Variables">
    <object class="Variable" serializationversion="1">
      <property name="name" class="String">checkinRecord</property>
      <property name="global" class="Boolean">true</property>
      <property name="initialAssignment" class="InitialVariableAssignment">
        <property name="type" class="TypeReference" serializationversion="0">
          <property name="typeName" class="String">CheckinRecord</property>
        </property>
      </property>
    </object>

    <object class="Variable" serializationversion="1">
      <property name="name" class="String" id="0">excelData</property>
      <property name="initialAssignment" class="InitialVariableAssignment">
        <property name="type" class="SimpleTypeReference">
          <property name="simpleTypeId" class="Integer">150</property>
        </property>
      </property>
    </object>

    <object class="Variable" serializationversion="1">
      <property name="name" class="String">dayNumber</property>
      <property name="initialAssignment" class="InitialVariableAssignment">
        <property name="type" class="SimpleTypeReference" id="1">
          <property name="simpleTypeId" class="Integer">7</property>
        </property>
      </property>
    </object>

    <object class="Variable" serializationversion="1">
      <property name="name" class="String" id="2">tempNV</property>
      <property name="initialAssignment" class="InitialVariableAssignment">
        <property name="type" class="SimpleTypeReference" id="3">
          <property name="simpleTypeId" class="Integer">12</property>
        </property>
      </property>
    </object>

    <object class="Variable" serializationversion="1">
      <property name="name" class="String">dateText</property>
      <property name="initialAssignment" class="InitialVariableAssignment">
        <property name="type" idref="3"/>
      </property>
    </object>

    <object class="Variable" serializationversion="1">
      <property name="name" class="String">p1</property>
      <property name="initialAssignment" class="InitialVariableAssignment">
        <property name="type" idref="1"/>
      </property>
    </object>

    <object class="Variable" serializationversion="1">
      <property name="name" class="String">p2</property>
      <property name="initialAssignment" class="InitialVariableAssignment">
        <property name="type" idref="1"/>
      </property>
    </object>
  </property>

  <property name="proxyServerConfiguration" class="ProxyServerConfiguration" serializationversion="0"/>
  <property name="httpClientType" class="HttpClientType">
    <property name="enum-name" class="String">WEBKIT</property>
  </property>
  <property name="ntlmAuthentication" class="NTLMAuthenticationType">
    <property name="enum-name" class="String">STANDARD</property>
  </property>
  <property name="usePre96DefaultWaiting" class="Boolean" id="4">false</property>
  <property name="maxWaitForTimeout" class="Integer">10000</property>
  <property name="waitRealTime" idref="4"/>
  <property name="privateHTTPCacheEnabled" class="Boolean" id="5">true</property>
  <property name="privateHTTPCacheSize" class="Integer">2048</property>
  <property name="comment"><null/></property>
  <property name="executionMode" class="ExecutionMode">
    <property name="enum-name" class="String">FULL</property>
  </property>
  <property name="avoidExternalReExecution" idref="4"/>

  <property name="transitionGraph" class="Body">
    <blockBeginStep class="BlockBeginStep" id="6"/>
    <steps class="ArrayList">

      <object class="Transition" serializationversion="3" id="7">
        <property name="name" class="String">Load File</property>
        <property name="stepAction" class="LoadFile">
          <property name="fileNameExpression" class="kapow.robot.plugin.common.support.expression.stringexpr.FileNameValueStringExpression">
            <property name="value" class="String">C:\Users\huyhh1\workspace\kofax\cham-cong\Library\Data\Tạo dữ liệu tính lương.xlsx</property>
          </property>
          <property name="output" class="kapow.robot.plugin.common.stateprocessor.rest.ToVariableOutputSpecification" serializationversion="1">
            <property name="variable" class="kapow.robot.plugin.common.support.AttributeName2">
              <property name="name" idref="0"/>
            </property>
          </property>
          <property name="browserConfigurationSpecification" class="BrowserConfigurationSpecificationWebKit" serializationversion="25">
            <property name="ancestorProvider" class="BrowserConfigurationSpecificationAncestorProviderForStep"/>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders" id="8"/>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <object class="Transition" serializationversion="3" id="9">
        <property name="name" class="String">Open Excel Data</property>
        <property name="stepAction" class="OpenVariable">
          <property name="variable" class="kapow.robot.plugin.common.support.AttributeName2">
            <property name="name" idref="0"/>
          </property>
        </property>
        <property name="elementFinders" idref="8"/>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <!-- LOOP Checkin -->
      <object class="Transition" serializationversion="3" id="10">
        <property name="name" class="String">Loop Checkin Rows</property>
        <property name="stepAction" class="LoopInExcel">
          <property name="loopDirection" class="LoopInExcel$LoopDirection" id="11">
            <property name="enum-name" class="String">ROWS</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders">
          <object class="ExcelElementFinder">
            <property name="detail" class="SpecifiedRangeCellFinderDetail">
              <property name="range" class="kapow.robot.plugin.common.support.expression.stringexpr.ValueStringExpression">
                <property name="value" class="String">Checkin!A2:C2000</property>
              </property>
            </property>
          </object>
        </property>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <object class="Transition" serializationversion="3" id="12">
        <property name="name" class="String">Extract Ma NV</property>
        <property name="stepAction" class="ExtractCell">
          <property name="attributeName" class="kapow.robot.plugin.common.support.AttributeName2">
            <property name="name" class="String">checkinRecord.MaNV</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders">
          <object class="ExcelElementFinder">
            <property name="detail" class="NamedRangeCellFinderDetail">
              <property name="rangeName" class="ElementName"><property name="name" class="String">1</property></property>
              <property name="usage" class="ColumnFromRange"/>
            </property>
          </object>
        </property>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <object class="Transition" serializationversion="3" id="13">
        <property name="name" class="String">Extract Ngay Lam Viec</property>
        <property name="stepAction" class="ExtractCell">
          <property name="attributeName" class="kapow.robot.plugin.common.support.AttributeName2">
            <property name="name" class="String">checkinRecord.NgayLamViec</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders">
          <object class="ExcelElementFinder">
            <property name="detail" class="NamedRangeCellFinderDetail">
              <property name="rangeName" class="ElementName"><property name="name" class="String">1</property></property>
              <property name="usage" class="ColumnFromRange">
                <property name="columnId" class="ByIndexExcelColumnId">
                  <property name="offset" class="kapow.robot.plugin.common.support.expression.integerexpr.ValueIntegerExpression">
                    <property name="value" class="Integer">1</property>
                  </property>
                </property>
              </property>
            </property>
          </object>
        </property>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <object class="Transition" serializationversion="3" id="14">
        <property name="name" class="String">Extract Gio Lam Viec</property>
        <property name="stepAction" class="ExtractCell">
          <property name="attributeName" class="kapow.robot.plugin.common.support.AttributeName2">
            <property name="name" class="String">checkinRecord.GioLamViec</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders">
          <object class="ExcelElementFinder">
            <property name="detail" class="NamedRangeCellFinderDetail">
              <property name="rangeName" class="ElementName"><property name="name" class="String">1</property></property>
              <property name="usage" class="ColumnFromRange">
                <property name="columnId" class="ByIndexExcelColumnId">
                  <property name="offset" class="kapow.robot.plugin.common.support.expression.integerexpr.ValueIntegerExpression">
                    <property name="value" class="Integer">2</property>
                  </property>
                </property>
              </property>
            </property>
          </object>
        </property>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <!-- Robust day parsing: between two "/" -->
      <object class="Transition" serializationversion="3" id="15">
        <property name="name" class="String">Normalize Date Text</property>
        <property name="stepAction" class="AssignVariable" serializationversion="4">
          <property name="stringExpr" class="Expression" serializationversion="1">
            <property name="text" class="String">removeSpaces(trim(checkinRecord.NgayLamViec))</property>
          </property>
          <property name="variable" class="kapow.robot.plugin.common.support.AttributeName2">
            <property name="name" class="String">dateText</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders"/>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <object class="Transition" serializationversion="3" id="16">
        <property name="name" class="String">Find Slash 1</property>
        <property name="stepAction" class="AssignVariable" serializationversion="4">
          <property name="stringExpr" class="Expression" serializationversion="1">
            <property name="text" class="String">indexOf(dateText, "/")</property>
          </property>
          <property name="variable" class="kapow.robot.plugin.common.support.AttributeName2">
            <property name="name" class="String">p1</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders"/>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <object class="Transition" serializationversion="3" id="17">
        <property name="name" class="String">Find Slash 2</property>
        <property name="stepAction" class="AssignVariable" serializationversion="4">
          <property name="stringExpr" class="Expression" serializationversion="1">
            <property name="text" class="String">p1 + 1 + indexOf(substring(dateText, p1 + 1), "/")</property>
          </property>
          <property name="variable" class="kapow.robot.plugin.common.support.AttributeName2">
            <property name="name" class="String">p2</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders"/>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <object class="Transition" serializationversion="3" id="18">
        <property name="name" class="String">Assign Day Number</property>
        <property name="stepAction" class="AssignVariable" serializationversion="4">
          <property name="stringExpr" class="Expression" serializationversion="1">
            <property name="text" class="String">(p1 &gt;= 0 &amp;&amp; p2 &gt; p1) ? toInteger(substring(dateText, p1 + 1, p2)) : -1</property>
          </property>
          <property name="variable" class="kapow.robot.plugin.common.support.AttributeName2">
            <property name="name" class="String">dayNumber</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders"/>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <!-- If invalid day => Next Iteration of Checkin loop -->
      <object class="Transition" serializationversion="3" id="19">
        <property name="name" class="String">Validate Day</property>
        <property name="stepAction" class="TestValue" serializationversion="0">
          <property name="condition" class="Expression" serializationversion="1">
            <property name="text" class="String">dayNumber &gt; 0 &amp;&amp; dayNumber &lt;= 31</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders"/>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0">
          <property name="reportingViaAPI" class="Boolean">false</property>
          <property name="reportingViaLog" class="Boolean">false</property>
          <property name="controlFlow" class="kapow.robot.robomaker.robot.ControlFlow$NextIteration"/>
        </property>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <!-- LOOP Timesheet -->
      <object class="Transition" serializationversion="3" id="20">
        <property name="name" class="String">Loop Timesheet Rows</property>
        <property name="stepAction" class="LoopInExcel">
          <property name="loopDirection" class="LoopInExcel$LoopDirection">
            <property name="enum-name" class="String">ROWS</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders">
          <object class="ExcelElementFinder">
            <property name="detail" class="SpecifiedRangeCellFinderDetail">
              <property name="range" class="kapow.robot.plugin.common.support.expression.stringexpr.ValueStringExpression">
                <property name="value" class="String">'Chấm công'!A3:AF2000</property>
              </property>
            </property>
          </object>
        </property>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <object class="Transition" serializationversion="3" id="21">
        <property name="name" class="String">Extract Temp NV</property>
        <property name="stepAction" class="ExtractCell">
          <property name="attributeName" class="kapow.robot.plugin.common.support.AttributeName2">
            <property name="name" idref="2"/>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders">
          <object class="ExcelElementFinder">
            <property name="detail" class="NamedRangeCellFinderDetail">
              <property name="rangeName" class="ElementName"><property name="name" class="String">2</property></property>
              <property name="usage" class="ColumnFromRange"/>
            </property>
          </object>
        </property>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <!-- If not match => Next Iteration of timesheet loop -->
      <object class="Transition" serializationversion="3" id="22">
        <property name="name" class="String">Match Employee</property>
        <property name="stepAction" class="TestValue" serializationversion="0">
          <property name="condition" class="Expression" serializationversion="1">
            <property name="text" class="String">trim(tempNV) == trim(checkinRecord.MaNV)</property>
          </property>
        </property>
        <property name="elementFinders" class="ElementFinders"/>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0">
          <property name="reportingViaAPI" class="Boolean">false</property>
          <property name="reportingViaLog" class="Boolean">false</property>
          <property name="controlFlow" class="kapow.robot.robomaker.robot.ControlFlow$NextIteration"/>
        </property>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <!-- Step exists, but content/offset MUST be set via DS UI to avoid version-specific serialization classes -->
      <object class="Transition" serializationversion="3" id="23">
        <property name="name" class="String">Set Timesheet Cell</property>
        <property name="stepAction" class="SetContentOfCell" serializationversion="0"/>
        <property name="elementFinders" class="ElementFinders"/>
        <property name="errorHandler" class="ErrorHandler" serializationversion="0"/>
        <property name="comment"><null/></property>
        <property name="enabled" idref="5"/>
        <property name="changedProperties" class="java.util.HashSet"/>
      </object>

      <object class="End" id="24"/>

    </steps>
    <blockEndStep class="BlockEndStep"/>
    <edges class="ArrayList">
      <object class="TransitionEdge"><from idref="6"/><to idref="7"/></object>
      <object class="TransitionEdge"><from idref="7"/><to idref="9"/></object>
      <object class="TransitionEdge"><from idref="9"/><to idref="10"/></object>
      <object class="TransitionEdge"><from idref="10"/><to idref="12"/></object>
      <object class="TransitionEdge"><from idref="12"/><to idref="13"/></object>
      <object class="TransitionEdge"><from idref="13"/><to idref="14"/></object>
      <object class="TransitionEdge"><from idref="14"/><to idref="15"/></object>
      <object class="TransitionEdge"><from idref="15"/><to idref="16"/></object>
      <object class="TransitionEdge"><from idref="16"/><to idref="17"/></object>
      <object class="TransitionEdge"><from idref="17"/><to idref="18"/></object>
      <object class="TransitionEdge"><from idref="18"/><to idref="19"/></object>
      <object class="TransitionEdge"><from idref="19"/><to idref="20"/></object>
      <object class="TransitionEdge"><from idref="20"/><to idref="21"/></object>
      <object class="TransitionEdge"><from idref="21"/><to idref="22"/></object>
      <object class="TransitionEdge"><from idref="22"/><to idref="23"/></object>
      <object class="TransitionEdge"><from idref="23"/><to idref="24"/></object>
    </edges>
  </property>

  <property name="browserConfigurationSpecification" class="BrowserConfigurationSpecificationWebKit" serializationversion="25"/>
</object>