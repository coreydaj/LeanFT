<?xml version="1.0" encoding="utf-8"?>
<applicationModel>
  <name>AdVantage Application Model</name>
  <className>VantModel</className>
  <description>Application Model For The AppVantage Web site</description>
  <objects>
    <object technology="Web" type="WebElement">
      <name>Amount</name>
      <codeName>AmountTransfered</codeName>
      <identification>
        <properties>
          <property name="TagName" type="string">DIV</property>
          <property name="InnerText" type="string" isRegExp="true">\Amount:.*</property>
        </properties>
      </identification>
      <children />
    </object>
    <object technology="Web" type="Page">
      <name>Advantage Online Banking</name>
      <codeName>AdvantageOnlineBankingPage</codeName>
      <identification>
        <properties />
      </identification>
      <children>
        <object technology="Web" type="EditField">
          <name>Password</name>
          <codeName>Password</codeName>
          <image>iVBORw0KGgoAAAANSUhEUgAAAJ8AAAAyCAIAAADN+56HAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAFTElEQVR4nO2cX0hbVxzHv+dyX/cyZ1OVhkDuNYWuDy2UrsYmdq8ZGzVgiqWw4gY1T7OFpAWHL4IaMGZQ0j209aEIzYPJNshrTUxC2+1hjFao5hZE8G/bMSbTdlN+e7j33FyltoOJrYfzedB8fucm94c/7znHc8+VrW++AgAwgFBDqgjK1jdfETcTqcKoAkfIWXqpAihb23wJgIHBaiOpwihbs8ZdiYCojmqblWdShVEFIAYAjHfZUsVR9tfmSx4EUBy4ksrDwdFvyheDvBVgcBy8V1oZac2iP9nT9m7T2IeqwoZHQ+HcNb/5sjo2FBuoBK/5He/ZevBeKHtjq9SdVcWb0JsO4N7zKqCjMtI6PmlGefmLA5etC/31EVemK71wIdnTBhRGz/bC7AaKA5d/cEdHzmvVsaHYjWXH241MV/oBXLPTy4H+ZE+bkelK350GQkcDVjqVkdZy42g8or8xaYmNAqrNo62Ou6bF+4+g1ekwMl3jc93RXHk4Nxr05MuZGaAwmjKCifJwrhw9Z5iR2ykjmCgnc6XoOaOcqXqPf+qamzdAqM6veHy0VCXAWDJcn5zQULgdu4Fzo8lcKRzIj/eMGea5Z7XWXCnZ00bFgfRdBBOlZO4UJq2E/D2leERz5Gp9l7qDqsSIO2OMAOSzZ/NZq/i+4NBNPxE6bg5HGCMCtDo3QPYvAoHg7bgZYwy0wAAQEcEbuRUnAubrZ+9MzXR6F+cOfHEBP/78tIOmHuBwtBkzvzxDyN+hEaHls+5y7N7UTOcRAjxuFxiIni4aCFz4XAOh7XTE92gRIOJJEoMzZ6k7qQri/TWZPz6EwtmrfvD6Wa3F22d7H/OIKwJQ8Muh+UT8dBGApzua7PTiNRFfoHdmAZVfjQ9DYeD+Mpp+n9V0jag4t8zHVNKa6vknw93kJSJgaXHaPLWV1PYkpf4XVVArNYe2lpaAcqr3caB/OFtKZkvtfBSE3hnPloazpejJe+lUEQDTO2OOCANajoVWloov5rQ6Xf/IbTzP3H8UOOUH0Oh22Seqzj/Ddg42+PC2rKS+TRXiHbPzgFqEanECiKpjlUkAhOpYon2wzFtdDQ2ojg21D1YcEQKh8RAe3nmCQwdBLce0Jw+Nj48FAILeWI98JTMDUCV/Y9lz5ohOzlN7j59xTd75qUpAYTIzjS1JbstZ6k6qMhAYiByzaUb25Npq9YcuVeLfXpkEPN3hiG98cQF6Zzjydbo9kAXguRRN6kBzOPKVI9IMAtNOHMb3T07GvGBocGMWzUHzk9suDnYnrnZdyQAItWfPewlPHSlAOx+OTKTjgSJ8wUgIi4wRyt8FKg23YpHmrTkzqTsoW91Yh0RQ7L93zX5x52UPqftQ2erGGhj4xBQMkCqMqgRedSDZn4Bkn9PX17e6aY2221ciC4XCXqcj2T2upwadqgIwV54Y77InJib2Pi3JrjD120M4CqrwRUVmry5KhIARyN41J0srGATnnkiJeMjqioxKtWVH2TkLAi8kqVtUIghWQRWASJZXLOyCKnbHLAssDHZBFdR2P0gEwS6oOWcm+4tECKyCKrKmAqPwXThUuzMo2efYBVXJDsiuWSQIgLkn0laJQBCgWhMsJqsrEmQW1Nkzy4FXEPgwy5RtQYlA0PadN9dTg+b9fYkAsOf//GlLnfrBO0xFslu82Fg1XzjvETEefY925Er9P6oQ+EIGEaxHUKQKokrtBiBjBJIqklqzKiLwy1mqOMqf37WRKpCqIDK7afPfKoDfT5AqgKrEzImWeWHbDyRIFUHZyt9/AO/PHF7qbqq6JQRIFUn/BdmIX56QWjGAAAAAAElFTkSuQmCC</image>
          <identification>
            <properties>
              <property name="Type" type="string">password</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">j_password</property>
            </properties>
          </identification>
          <tags>
            <tag>EditField</tag>
            <tag>Encrypted</tag>
          </tags>
          <children />
        </object>
        <object technology="Web" type="EditField">
          <name>Username</name>
          <codeName>Username</codeName>
          <image>iVBORw0KGgoAAAANSUhEUgAAAKAAAAAqCAIAAAC7hsfEAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAEUElEQVR4nO2bT2gbRxTGvyeW0ntLBBKlNcVrDM0p9iEoBUnXChlaHdI19sE9pilR8cE5xKZuStxaSQxxjk0hxnvyJcY6FSyB6paCLiUUE9XFObhNC/G1TdNIr4f5o11ZocUWDbu872D5NzM7O+ibN7Mzs6JOh0EAGACYAAjGCanT6ZgknSgYJ6Q2d4h1kv4UjBFSW0UwQACjtliuoFSdzyhE696H3uF7zZkcBTqHKfyfcadyppHy5zwXtcVyZTOQVSxVFzInqFnw39DRsRzOs32hP+JYaG5ULFXnzwEA9vyp1Xc+gcFj1yz4fHTAJp6DWdyDQCD4iqXqfAZQ8V3fB4DTs82ZHPDg7tLHX2No9/f9Yqk6nwnXzNxtAwNgvPn+ley3XsM/n/HcneWxRsqfm3RDzbPNEDwmOjqROBTiR3D7y8pP2RvNCRd7/tTGeuvcpPvNsld//Yubt/JUW7xUmbqXWpsAgN1T7zYv5023sBNCuHuRznXfOjta//WA2M3MNjOB+3K4GYLHxQTImM0mD/1QGwMMe2uXPZd5+0Edp8/kCUDufHZo9/AXVc/oK2l9ra1ZVRREfj6y4CDRAUMndCOWdDk7ejPlP7h5cK08Vgcw9NGF1elhAMD962OXrmvvk49anAYw/OqIvtbWrP6G0OT+9mi3J7fnWsGToQP1ZetgSr+RxEOwnYIPDvcBZmJgeGpuaxrgPX/69nJ6ZRbAaPbG3QnX9BcQWt/pnhOwkWA7jhkT2I4mrccPkTybtrnmviQ4IEzY754YIDd9CpsNv6XwZ/+r+yiO5MCttWuFT3eMUcnUa4z8SHa3vrkNqNzxOzU7oqvxv1uz/gxI5e75V+v7xbc9N1i451rBk6F5yFIGMPIzW5/fKUyWfZVYLG1dyQBwp0re9O3C+AaAoYsXbg0DyMyuP744WS4AQNJbn8kBrWBVoZoD2twobG4gVD8DO5XxRmp9znOPXit4AqRnnTZE8ZUT+N9OnYLxQcc8WTOFijAEo4zWUMf6ysBni1chirgWFhba3GHjtMNAcGqu1+svpFmigWh1ZQmA3aIiIudooVqt9r82SjQ4/fjD90Fk5kTfckTUN10UMZHd6BDFUcxw6MhGhCjqChrq9CydRPGQXRw5AMOc2YpiI7saTgRQXI6NuobahyxmGadjJVbnSQliuzCWCI6LjKHEcEJ70KJYiM3SiM15sBqeZYiOjbqGJgIPXBLF8ZJ66U4WSTFU98iQHWIZmeMn7SmxPvAHADlfiI30q7Okd7JMqgzU8ZIy1On77MzidqQV8NRhE9AqdXVlqefQWBQ5sfplAQFg+rP9FKTj+eXESy+6baIB6An/rQwlkGOnZID/4qc6HQCYCILRRG0oq9MkuxttziD0K3mCkUXYxVH3nSxGSIIxQAboj/YT2BkZAJjMQ5hgRBEBQxMAqc1K5T+Ff5EvGEUMGtp98Z0A82NeNlO0YHRRGYp/ALhrRWzw0ZYfAAAAAElFTkSuQmCC</image>
          <identification>
            <properties>
              <property name="Type" type="string">text</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">j_username</property>
            </properties>
          </identification>
          <tags>
            <tag>EditField</tag>
          </tags>
          <children />
        </object>
        <object technology="Web" type="Button">
          <name>Login</name>
          <codeName>Login</codeName>
          <image>iVBORw0KGgoAAAANSUhEUgAAADoAAAAnCAIAAADRkcw6AAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAA7EAAAOxAGVKw4bAAACuElEQVRYhdWYsUtbURTGf/dxWzBbarCFKMFSRDIIZnDJEDK8oaAg5A/QoYOzi+AfIHRxdnCodOhQUwLN0PYtZnCog6VCg0SpPGwG09h20rbL6XDz0tTGmvi0uX5DyHeSc+95h3PvOe9TRz+/AShRIKIwsJVqAARBaMJeqlG/rUpJYLeUakQByvwUpN1aqiVgEnzYTHWryRhtpholIty7FcVuZLPZZ29eaBFl8vzh/dseR3Q+1guvzBcNIv/+r01weh1Ad3BuUGoF0cjNKQbBQQV32kXYKni5gh9mO7+0mVvdvfwSCkf+6MtWQwQN/Ld4E5l0PhPCXwUjTihUdxfWDvcBIrMz6am4sforSxUPIOYm6x4j+emEX9qc3+tffjRKaXN+L+JS92oA7qQ7N9bBRoIT+qD5K2uHw5NuftFdTvNk7d0WAFuFipccyS+6+Zm+g3I7v1qdCeMV8YoNrwvjdQhZCztfPGKpMYBEZtClvr0D+Ntl3PsJgPhoLtnOcWDoofGKRjrdSymns1vhXPhfTxjou3vGWj09IBLvD7f0XxCRsG0iEY1QOz06Y433DXNSPQ63dDuELoaxO0EB4Jc+BYWRSCXxPvoA1d1829rtHmLG8y48ypVcuRIQcw8k5mZOF9a8XNFYxicAmJgecZcquXKFgaHZJPtXEzDq8MdnYPB27PoGyK2Cl4+mHmcuX8vrhVcbGxtPXz935Fq6xPHLVW9lB2gUw3D0Cs6debXsdGboBv1TE7Fc0fOKAA/SqY66QAfQ1xArAGPj+SsKsRU3bDzXImIS3HwfshiiFYiQzWZ7HcrFEEH532u9DqMLaBCUUflaxB1bqTYyg1GjGnKZxVQjAqpx3MQ0DXupA2AUyWZ3s5g6gaWlE1tMdYvQ13gKm6nTYm1RVG2lTX03kKrtprr5FK2pt5b+AklgVSoeLsckAAAAAElFTkSuQmCC</image>
          <identification>
            <properties>
              <property name="ButtonType" type="string">submit</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">Login</property>
            </properties>
          </identification>
          <tags>
            <tag>Button</tag>
          </tags>
          <children />
        </object>
        <object technology="Web" type="Link">
          <name>Logout</name>
          <codeName>Logout</codeName>
          <image>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</image>
          <identification>
            <properties>
              <property name="TagName" type="string">A</property>
              <property name="InnerText" type="string">Logout</property>
            </properties>
          </identification>
          <tags>
            <tag>Link</tag>
          </tags>
          <children />
        </object>
        <object technology="Web" type="ListBox">
          <name>fromAccount</name>
          <codeName>FromAccount</codeName>
          <image>iVBORw0KGgoAAAANSUhEUgAAAFwAAAAWCAIAAADCTmwFAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAA7EAAAOxAGVKw4bAAADJ0lEQVRYhe1VPXPbRhB9u/LvYGZMqhFrVSdkTKV2/kJUkGni4og+KZLeZqMGAGecvyDXBDQToXJtN/go5L8R76a4A0AgJG2VHmMb8u17t/t2eQDpX/0MAEoEVQCEEdJnFZd2nBd935DVYzTkCP1NUYCgCgJAI3zmVtNgjBAgEhVVELV3iEbIAEBQQEHNyxdlEjCdMZ0x8RkzE6+zjh2IvwlYJgGv779S/KzZkw8HCYCJiofl80PstwjpKWImItUufQpm4VkQBMR8lVQOMnMLgfuQg3gbEjMxr7N6GzAxU+Do/VJVEvize2wZOz0HSeXEVeIzP25rdzZbs7u8QVx7TxRsa1IlVElAYQbQfchXYXjl629r0jQ8X+XY/MTrjOqtk52aV1REVFRFRFQcLGIDExUN9Gxq4bKqUkQGsKmIlJEB7E5UdhaA3TUs7E5Ed9azXSmXUumxqYWJC1WRnYWJSvcZF15lokKL2PiDrme6JxaVMjaOTS08W3hvbqKBjb15B5ABArlnyj1UDgL5r+fuN2cKkrK5ZJfTmQKKTx9zE/+2INDz5V8Wm3eZu6H25QKK2ewSJnp1DdLJ3PiybaPrN6JvFqB9NrvbmJsXUyjo+rU+LKeUvduYmxdTIuD6tT4sZ/r4ITfxqwUI09WfFpu7tHHVVHbfvQ8FzRY3pptoYIOOQvb/0mjfMx7CxKWqqIjqP8upZ83FBASgrt6jFU8uTOPFzH9ofXWl+pUJ1faKmIiJz1e5y1Tv0bdRug57ruo2A7htHmsEmPnk/xMNbRyFjC4UvTgFp9PLDjx+zIET4gGskl9WiAsVUfdEaL+e9ju0Z13Gw8cP+bDyE22cgAxQn/46OLkw+eo2I6BMft/Avlwc6PTFUuntyg8ymZv8bVoBQLYmCjOazE3+NnOZkDjMdDI3+eo2A1Bt/9jA/ux65n9nNaDZ3ebIwE8fUFW0F6KqZWzc49NjUwsTlfvQhVdqagdfm1I21V6p9qRNuzPSryf9DsWRnl3Kxk2nPR/+5dvKbKpl1Brqu+ogqQ5u0RjgL0u+vxiXciDGpRyIcSkHYlzKgRiXciD+A8Zy5F2COeQWAAAAAElFTkSuQmCC</image>
          <identification>
            <properties>
              <property name="TagName" type="string">SELECT</property>
              <property name="Name" type="string">fromAccount</property>
            </properties>
          </identification>
          <tags>
            <tag>SelectionBox</tag>
          </tags>
          <children />
        </object>
        <object technology="Web" type="ListBox">
          <name>toAccount</name>
          <codeName>ToAccount</codeName>
          <image>iVBORw0KGgoAAAANSUhEUgAAAEYAAAAYCAIAAADIyBxUAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAA7EAAAOxAGVKw4bAAABg0lEQVRYhe2XPXKDMBCFHzkIjdxRpxINOUCuEBfQpZAOkCoHMIU7KHIGHwA1UKV2ZzW+yKYICGHxl4knDh6+Snq72tnVip3BIyLcFw+3TuD6rCUtgbWkJXDHJek89C4Jc33T3PrReehJNeLQlMTikoiIThkHz05ERFTG7C9yvDYzHp6SI12zutuajWgkV3GiKmmWphFKeqGU9lElN0mF9MmTarBf1MXuUrOHKDqrlkIYqV0WoonQrBylL2rrRKeMG83xM8Yhprp0PlY8e40AsPhdID10biXaEe0iAIAf8Pr6DynfRqw2lzHrUcajdhHPEQAWbflEqg0TJWn92W5M1pbdvLtNUrkn5ig9UW144I+n6DBREmOP7eZ8rC5Sy18SWI/FOTFHcaL2Sj9g6uH5Aa+SvQKg87e0fgU9qH3dJfgBrz6UBr5HgFQDSk/U2ksd0l9UNDkemm8UwKXesYnC+rzdE3MUI4nMGg/OyKjdRDE0KDxa/5f+P2tJS2AtaQmsJS2BLyxNvbnJW/niAAAAAElFTkSuQmCC</image>
          <identification>
            <properties>
              <property name="TagName" type="string">SELECT</property>
              <property name="Name" type="string">toAccount</property>
            </properties>
          </identification>
          <tags>
            <tag>SelectionBox</tag>
          </tags>
          <children />
        </object>
        <object technology="Web" type="Button">
          <name>Next</name>
          <codeName>NextButton</codeName>
          <image>iVBORw0KGgoAAAANSUhEUgAAADgAAAAkCAIAAABT8G6pAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAA7EAAAOxAGVKw4bAAACAElEQVRYhe2XQWvaYBjH/xn7FpIv0IWRERQvg3gsdFLMJVQCUtAdcvEiQg4JWihILsLoSfAiFmEzFCnYY45dJOjGuttOIdBv8fbgsE0TN2OCGsjvlDxJeH/88z5v3lCEECSBN/sW2JRUNG5S0bhJjOhbf4miqN17rGO1egYnSg4AnudfKgUkuuThx/f4cgnNt5u7V5XEzNFUNG62E3XHlTxTmTirwqzrOd3g8c4s3JAREl1cXoccLArbi0plcVDtWr66Y9QYNs+weUa3AU/YjlFjdNvSBW2BQTVcqBES5ev98uiL4XqKs+5x+6i/uP+1uJKGsmS4yNanxVvFcAH7enIybXC5xrjFQurdN7O7EQVyooJ2a/xiblrmCOWPOQDgztT38z+PAOiSdjppSRUZcpHedqxoXU8XL1VoI9tTHMrLV3/c/okH1wGAjCAfzaGchYnwFWu/TBtClz5LrNyBuKp8UMeDUsZ7l92poqXeKgbnu7Qp0ddRrtkTB8PR8iTHi/P2Vwv4u4TpNgBLl9GrCyXtdOKZJzsWBbL1fvn5eKr+PmfzDCtoUKYNzjFq57hqZgFkhIuTm09dCxn6Xeiup4jvL5SiKELI3jclpmmapkn+vc07QFLRuFm7PPm3rvslWLRQKOzY478Ed/1eVAJZ6QWIHiaJaaZUNG5S0bhJjOgTKV01XhHcqycAAAAASUVORK5CYII=</image>
          <identification>
            <properties>
              <property name="ButtonType" type="string">submit</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">Next</property>
            </properties>
          </identification>
          <tags>
            <tag>Button</tag>
          </tags>
          <children />
        </object>
        <object technology="Web" type="EditField">
          <name>amount</name>
          <codeName>Amount</codeName>
          <image>iVBORw0KGgoAAAANSUhEUgAAAD0AAAAZCAIAAAANPbGeAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAA7EAAAOxAGVKw4bAAABdUlEQVRYhe2XLXLDMBBGv+1BQhQWbKQQ5wA9Q0DMCjYHCOoBIhAmg5zBB5BIhYrDLJKLqMCu7WTin2Y66nTGD1lr7erZWmvGFELAP+TlrwWeZPaOy+wdl9k7LmPePl8T7W0MFZ+vp6804m1PWcKsiijiPyIMYRhsgmFIXbYhqTVXyWxKLQHgdgJuQp30UkuwqWPMsjOxSWPTTutn0LvSDqHUsvEyXNWuJNBodGJ3Vz3eDzNGbFsG+sTnB8WvKQCRbpGd2l6poyKB1G8pgMWqfnXXi6tDYvfOGGywprac3h0N/d7enh3UhoiIlpnrOMjVoi/Hf7aD5mEe019lCr3e3p5dZ9dKLSd8nUIk7eB6cff3H4SepM/bnjJXbWRtlG6lOuR+pNxiJV3VUT4/KHxXcGfrAdhC/YYz0Otti3bRCpFuZb38AGL3YVhtiGiZQZfHFEB6NOyyJREVKz3QOkIkUBva2yknOYX5vyEms3dcZu+4zN5x+QL5eXB9OCvOFQAAAABJRU5ErkJggg==</image>
          <identification>
            <properties>
              <property name="Type" type="string">text</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">amount</property>
            </properties>
          </identification>
          <tags>
            <tag>Money</tag>
            <tag>EditField</tag>
          </tags>
          <children />
        </object>
        <object technology="Web" type="EditField">
          <name>transferDate</name>
          <codeName>TransferDate</codeName>
          <image>iVBORw0KGgoAAAANSUhEUgAAACYAAAAbCAIAAACfu2qKAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAA7EAAAOxAGVKw4bAAABDUlEQVRIie2WIZLDIBiFX3qQGuKqq4hJL9KKxq0gB6jaAzRiHRF7ElCo6rpgepF/RYDQZJvuZGc6FTwF+Xl8AR4zZESE12r1Yl5CJmRCvjnStkUWq2jtjM+2RVbrfyIBgMuOnNS2yuepi+UR1EkeI4mUAIQKtfivlHC9vh663t9J7q1TzSBJiQHgJgjNYdLB10mOx6Sgv8SnPBOdSwDAesNHRau/DT+UDAA7fgpc7LPDmEFae/GtkKy8Mr+MNFXel3cNzPW2HHm7GmwZg233FaJDmupuM91+LEDqetdw+XHv11+TVbLywJtTH21dP7tbAEaJjRVfmCGgPlLu2yixccIf5iij9BBJyIR8N+QPVis+YJd+TTYAAAAASUVORK5CYII=</image>
          <identification>
            <properties>
              <property name="Type" type="string">text</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">transferDate</property>
            </properties>
          </identification>
          <tags>
            <tag>Date</tag>
          </tags>
          <children />
        </object>
        <object technology="Web" type="Button">
          <name>OK</name>
          <codeName>OKButton</codeName>
          <image>iVBORw0KGgoAAAANSUhEUgAAADAAAAAoCAIAAAA35e4mAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAByElEQVRYhe2YMWvCQBSAX0p/REGydG0oikQCXZKx0DhUClkCpZAuLi6lILRBCi4uLi7NUsgitBaJBTtJoEsakVqwTm7B33EdtDHRxIuhxgz5yPKOx92Xu3fJJYACM1FOoNQLnh8OCJjXKwFEoIMQgRCCOLG3a4FlEiEciRCORAhH7IT2V5sIgojew34+e8/Q9t8QC1iWdQ7tMUMzfr4//+/+fXlpvy+1xK6GEiEc4YXMGkOlZ9d1y/pr7depdN0EZ8iIr9Pg3foW9VqmrcuCfNQYDbMAAJYm8sxEMW5pd1a/TknjSscokBt0HWqG+s8ylLs32XlI5lVFUBua5cyxNFEaVzqPG9mEFDL1ZiafdQ10cJgZvpm2kaWJfJVSNraBsEu2ApmiFkHzigcAoPyS17KVXSYqxkgRVMlR7IEJI5RjhS9t4Bqr/6Gmz3LzBRI4GoAude9BvtM2VQo1Q/RFBaqntcE8tDRRaorF/FLBkOeyKy0Y4WooVXgyyBpDpWfhsc/eThUeym2+SEFjZG9JHB7fZQRBIIQie7nquq7rOub4sUMSIRy+Rb16dIoGbyGO4yL2sPHeZdF72BrJ7xgciRCORAhH7IR+AbbAC7b1Of3wAAAAAElFTkSuQmCC</image>
          <identification>
            <properties>
              <property name="ButtonType" type="string">submit</property>
              <property name="TagName" type="string">INPUT</property>
              <property name="Name" type="string">OK</property>
            </properties>
          </identification>
          <tags>
            <tag>Button</tag>
          </tags>
          <children />
        </object>
      </children>
    </object>
  </objects>
</applicationModel>