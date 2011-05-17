<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" 
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                xmlns="http://schemas.datacontract.org/2004/07/ExchangeIntegration.Service"
                >
    <xsl:output method="xml" indent="yes" omit-xml-declaration="yes"/>

  <xsl:template match="/soap:Envelope/soap:Body/m:SendNotification/m:ResponseMessages/m:SendNotificationResponseMessage">
    <xsl:if test="m:Notification">
      <xsl:apply-templates select="m:Notification" />
    </xsl:if>
  </xsl:template>

  <xsl:template match="m:Notification">
    <SubscriptionEventNotification>
      <SubscriptionId>
        <xsl:value-of select="t:SubscriptionId"/>
      </SubscriptionId>
      <PreviousWatermark>
        <xsl:value-of select="t:PreviousWatermark"/>
      </PreviousWatermark>
      <MoreEvents>
        <xsl:value-of select="t:MoreEvents"/>
      </MoreEvents>
      <IsError>false</IsError>
      <Events>
        <xsl:apply-templates select="t:ModifiedEvent|t:CreatedEvent" />
      </Events>
    </SubscriptionEventNotification>
  </xsl:template>

  <xsl:template match="t:ModifiedEvent">
    <BaseExchangeEvent>
      <EventType>Modified</EventType>
      <xsl:call-template name="Ebase" />
    </BaseExchangeEvent>
  </xsl:template>

  <xsl:template match="t:CreatedEvent">
    <BaseExchangeEvent>
      <EventType>Created</EventType>
      <xsl:call-template name="Ebase" />
    </BaseExchangeEvent>
  </xsl:template>

  <xsl:template name="Ebase">
    <TimeStamp>
      <xsl:value-of select="t:TimeStamp" />
    </TimeStamp>
    <ItemId>
      <xsl:value-of select="t:ItemId/@Id"/>
    </ItemId>
    <FolderId>
      <xsl:value-of select="t:FolderId/@Id" />
    </FolderId>
    <Watermark>
      <xsl:value-of select="t:Watermark"/>
    </Watermark>
  </xsl:template>
  
</xsl:stylesheet>
