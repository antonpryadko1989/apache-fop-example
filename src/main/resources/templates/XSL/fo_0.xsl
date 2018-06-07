<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" version="2.0">
	<xsl:output encoding="UTF-8" method="xml" reverse-page="true" format="xml"/>
	<!--<xsl:output encoding="UTF-8" indent="yes" method="xml" standalone="no" omit-xml-declaration="no"/>-->
	<xsl:template match="body">
		<fo:root>
			<fo:layout-master-set>
				<fo:simple-page-master master-name="A4-portrail" page-height="297mm" page-width="210mm" margin-top="20mm" margin-bottom="25mm" margin-left="20mm" margin-right="20mm">
					<fo:region-body/> <!--margin-top="25mm" margin-bottom="20mm"-->
					<fo:region-before region-name="xsl-region-before" extent="25mm" display-align="before" precedence="true"/>
				</fo:simple-page-master>
			</fo:layout-master-set>	
			<fo:page-sequence master-reference="A4-portrail">
				<fo:flow flow-name="xsl-region-body" border-collapse="collapse" reference-orientation="0" font-family="Arial Unicode MS">
				<fo:table table-layout="fixed"  width="165mm" border-color="black" border-width="0.5mm" 
				border-style="solid" text-align="right" display-align="center" space-after="-1mm">
					<fo:table-body>
						<fo:table-row height="10mm">
							<fo:table-cell>
								<fo:block><fo:external-graphic src="E:work/1/image.bmp"  content-height="scale-to-fit" scaling="non-uniform"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
					</fo:table-body>	
				</fo:table>
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm" border-style="solid" text-align="center" display-align="center" space-after="1.5mm" >
					<fo:table-column column-width="125mm"/>
					<fo:table-column column-width="22mm"/>
					<fo:table-column column-width="18mm"/>
					<fo:table-body >
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" >
								<fo:block font-size="9.7px" font-weight="bold">Template Name /  אישור ביצוע פעילות מתקנת</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">Version</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">Date</fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="9.7px" font-weight="bold"><xsl:value-of select="templateName"/></fo:block>									
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold"><xsl:value-of select="version"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold"><xsl:value-of select="date"/></fo:block>
							</fo:table-cell>
						</fo:table-row>							
					</fo:table-body>
				</fo:table>
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm" border-style="solid" text-align="center" display-align="center" space-after="1.5mm" >
					<fo:table-column column-width="45mm"/>
					<fo:table-column column-width="40mm"/>
					<fo:table-column column-width="40mm"/>
					<fo:table-column column-width="40mm"/>
					<fo:table-body>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">Main Contractor</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="mainContractor"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">Project Name</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="projectName"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">Management Company</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block  font-size="5.4px"><xsl:value-of select="managementCompany"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">Contract No</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block  font-size="5.4px"><xsl:value-of select="contractNo"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">QC Company</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="qcCompany"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">QA Company</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="qaCompany"/></fo:block>
							</fo:table-cell>
						</fo:table-row>							
					</fo:table-body>
				</fo:table>					
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm" border-style="solid" text-align="center" display-align="center" space-after="1.5mm" >
					<fo:table-column column-width="85mm"/>
					<fo:table-column column-width="80mm"/>
					<fo:table-body>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">NCR Number</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrNumber"/></fo:block>
							</fo:table-cell>	
						</fo:table-row>
					</fo:table-body>
				</fo:table>
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm" 
				border-style="solid" text-align="center" display-align="center" space-after="-1mm">
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="39mm"/>
					<fo:table-body>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">NCR Opened</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">Position</fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">Name</fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">Date of NCR</fo:block>
							</fo:table-cell>	
						</fo:table-row>
					</fo:table-body>
				</fo:table>					
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm"
				border-style="solid" text-align="center" display-align="center" space-after="-1mm">
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="39mm"/>
					<fo:table-body>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-rows-spanned="2">>
								<fo:block font-size="6.5px" font-weight="bold">QA / QC</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">QCM</fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrOpenQCName"/></fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrOpenQCDate"/></fo:block>
							</fo:table-cell>	
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">QAM</fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrOpenQAName"/></fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrOpenQADate"/></fo:block>
							</fo:table-cell>	
						</fo:table-row>
					</fo:table-body>
				</fo:table>	
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm"
				border-style="solid" text-align="center" display-align="center" space-after="-1mm">
					<fo:table-body>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-style="solid">
								<fo:block font-size="9.7px" font-weight="bold">Details of Structure</fo:block>
							</fo:table-cell>
						</fo:table-row>
					</fo:table-body>
				</fo:table>
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm" border-style="solid" text-align="center" display-align="center" space-after="1.5mm">
					<fo:table-column column-width="21mm"/>
					<fo:table-column column-width="21mm"/>
					<fo:table-column column-width="21mm"/>
					<fo:table-column column-width="21mm"/>
					<fo:table-column column-width="21mm"/>
					<fo:table-column column-width="21mm"/>
					<fo:table-column column-width="24mm"/>
					<fo:table-column column-width="15mm"/>
					<fo:table-body>
						<fo:table-row height="7mm" font-size="6.5px" font-weight="bold">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Element (Station, Road, Tunnel, Bridge) / Element (Station, Road, Tunnel, Bridge)</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Stucture</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Element</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>From Section</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Till Section</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Side</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>NCR Level</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Number Of Days Late</fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm" font-size="5.4px">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block><xsl:value-of select="elementStationRoadTunnelBridge"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block><xsl:value-of select="stucture"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block><xsl:value-of select="element"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block><xsl:value-of select="fromSection"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block><xsl:value-of select="tillSection"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block><xsl:value-of select="side"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block><xsl:value-of select="ncrLevel"/></fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block><xsl:value-of select="numberOfDaysLate"/></fo:block>
							</fo:table-cell>			
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-columns-spanned="2">
								<fo:block font-size="6.5px" font-weight="bold">Expected Closing Date</fo:block>
							</fo:table-cell>							
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"  number-columns-spanned="6">
								<fo:block font-size="5.4px"><xsl:value-of select="expectedClosingDate"/></fo:block>
							</fo:table-cell>
						</fo:table-row>				
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-columns-spanned="2">
								<fo:block font-size="6.5px" font-weight="bold">Updated Expected Closing Date</fo:block>
							</fo:table-cell>							
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"  number-columns-spanned="6">
								<fo:block font-size="5.4px"><xsl:value-of select="updatedExpectedClosingDate"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-columns-spanned="2">
								<fo:block font-size="6.5px" font-weight="bold">Sub Project</fo:block>
							</fo:table-cell>							
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"  number-columns-spanned="6">
								<fo:block font-size="5.4px"><xsl:value-of select="subProject"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-columns-spanned="2">
								<fo:block font-size="6.5px" font-weight="bold">NCR Description</fo:block>
							</fo:table-cell>							
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"  number-columns-spanned="6">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrDescription"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-columns-spanned="2">
								<fo:block font-size="6.5px" font-weight="bold">Responsible Party</fo:block>
							</fo:table-cell>							
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"  number-columns-spanned="6">
								<fo:block font-size="5.4px"><xsl:value-of select="responsibleParty"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-columns-spanned="2">
								<fo:block font-size="6.5px" font-weight="bold">Corrective Action</fo:block>
							</fo:table-cell>							
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"  number-columns-spanned="6">
								<fo:block font-size="5.4px"><xsl:value-of select="correctiveAction"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-columns-spanned="2">
								<fo:block font-size="6.5px" font-weight="bold">Description Of Performed Corrective Action</fo:block>
							</fo:table-cell>							
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"  number-columns-spanned="6">
								<fo:block font-size="5.4px"><xsl:value-of select="descriptionOfPerformedCorrectiveAction"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-columns-spanned="2">
								<fo:block font-size="6.5px" font-weight="bold">Responsible Person</fo:block>
							</fo:table-cell>							
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"  number-columns-spanned="6">
								<fo:block font-size="5.4px"><xsl:value-of select="responsiblePerson"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-columns-spanned="2">
								<fo:block font-size="6.5px" font-weight="bold">Remarks</fo:block>
							</fo:table-cell>							
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"  number-columns-spanned="6">
								<fo:block font-size="5.4px"><xsl:value-of select="remarks"/></fo:block>
							</fo:table-cell>
						</fo:table-row>
					</fo:table-body>
				</fo:table>			
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm"
				border-style="solid" text-align="center" display-align="center" space-after="-1mm">
					<fo:table-body>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-style="solid">
								<fo:block font-size="9.7px" font-weight="bold">Acceptance of corrective action</fo:block>
							</fo:table-cell>
						</fo:table-row>
					</fo:table-body>
				</fo:table>
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm" 
				border-style="solid" text-align="center" display-align="center" space-after="-1mm">
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="39mm"/>
					<fo:table-body>
						<fo:table-row height="7mm" font-size="6.5px" font-weight="bold">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>NCR Closed</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Position</fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Name</fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Closing Date</fo:block>
							</fo:table-cell>	
						</fo:table-row>
					</fo:table-body>
				</fo:table>					
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm" border-style="solid" text-align="center" display-align="center" space-after="1.5mm">
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="39mm"/>
					<fo:table-body>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid" number-rows-spanned="2">>
								<fo:block font-size="6.5px" font-weight="bold">QA / QC</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">QCM</fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrCloseQCName"/></fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrCloseQCDate"/></fo:block>
							</fo:table-cell>	
						</fo:table-row>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="6.5px" font-weight="bold">QAM</fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrCloseQAName"/></fo:block>
							</fo:table-cell>	
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block font-size="5.4px"><xsl:value-of select="ncrCloseQADate"/></fo:block>
							</fo:table-cell>	
						</fo:table-row>
					</fo:table-body>
				</fo:table>
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm"
				border-style="solid" text-align="right" display-align="center" space-after="-1mm">
					<fo:table-body>
						<fo:table-row height="7mm">
							<fo:table-cell border-color="black" border-style="solid">
								<!--<fo:block font-size="9.7px" font-weight="bold">Additional Documents</fo:block>-->
								<fo:block font-size="9.7px" font-weight="bold">Additional Documents</fo:block>
							</fo:table-cell>
						</fo:table-row>
					</fo:table-body>
				</fo:table>
				<fo:table table-layout="fixed" width="165mm" border-color="black" border-width="0.5mm" border-style="solid" text-align="center" display-align="center" space-after="1.5mm">
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="42mm"/>
					<fo:table-column column-width="23mm"/>
					<fo:table-column column-width="19mm"/>
					<fo:table-column column-width="39mm"/>
					<fo:table-body>
						<fo:table-row height="7mm" font-size="6.5px" font-weight="bold">
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Item</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Exist/Does not exist</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Certificate No</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Expiration</fo:block>
							</fo:table-cell>
							<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block>Attached Documents</fo:block>
							</fo:table-cell>
						</fo:table-row >
						<xsl:if  test="not(additionalDocuments)">
						<fo:table-row height="7mm">
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell> 							
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell>								
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell>								
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell>								
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell>
							</fo:table-row>
							<fo:table-row height="7mm">
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell> 							
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell>								
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell>								
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell>								
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid"><fo:block></fo:block></fo:table-cell>
							</fo:table-row>--
						</xsl:if>
						<xsl:for-each select="additionalDocuments">
							<fo:table-row font-size="5.4px" height="7mm">
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="item"/></fo:block>
								</fo:table-cell> 							
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="exists"/></fo:block>
								</fo:table-cell>								
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="certificateNo"/></fo:block>
								</fo:table-cell>								
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="expiration"/></fo:block>
								</fo:table-cell>								
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="attachedDocuments"/></fo:block>
								</fo:table-cell>
							</fo:table-row>
						</xsl:for-each>		
					</fo:table-body>
				</fo:table>
			</fo:flow>
		</fo:page-sequence>
	</fo:root>
</xsl:template>
</xsl:stylesheet>
					
