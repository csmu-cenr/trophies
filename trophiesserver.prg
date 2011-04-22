DEFINE CLASS TrophiesServer as Custom OLEPublic

	ADD OBJECT DebugMessage as DebugMessage
	ADD OBJECT Queries as Queries
		
	DefaultPath = "g:\tbsdata"
	ExceptionMessage = ""
	
	XML = ""
	
	_recordNumber = 0
	_recordsAffected = 0
	
	* _queries = ""
	
	StartupPath = ""
		
	PROCEDURE INIT
	
		SET RESOURCE off
		SET EXCLUSIVE off
		SET REPROCESS TO 2 seconds
		SET CPDIALOG off
		SET DELETED on
		SET EXACT off
		SET SAFETY off
			
		this.StartupPath = ADDBS(JUSTPATH(application.ServerName ))
		SET PATH TO ( this.StartupPath)
		
	ENDPROC
	
	PROCEDURE AppendFromXML ( tableName_ as string, tableIDColumnName_ as string, cursorName_ as string, cursorIDColumnName_ as string, xml_ as String, debug_ as Boolean  ) as integer  
		
		LOCAL result as Integer 
		LOCAL xmlText as String 
		LOCAL selectCursor as String 
		LOCAL selectTable as String 
		LOCAL useTableNameInZero as String 
		LOCAL replaceID as String 
		LOCAL rowIndex as Integer 
		LOCAL idValue as Integer 
		LOCAL getIdValue as String 
		
		this.SetDefaultTo( this.DefaultPath )
			
		selectCursor = 'select ' + cursorName_
		selectTable = 'select ' + tableName_ 
		replaceID = 'replace ' + ALLTRIM(cursorName_) + '.' + ALLTRIM(cursorIDColumnName_) + ' with ' + ALLTRIM(tableName_) + '.' + ALLTRIM(tableIDColumnName_ )
		useTableNameInZero = 'use ' + ALLTRIM(tableName_) + ' in 0'
		IF ( ALLTRIM(tableIDColumnName_ ) == 'recno()' ) THEN 
			getIdValue = 'idValue = RECNO()'
		ELSE
			getIdValue = 'idValue = ' + ALLTRIM(tableName_) + '.' + ALLTRIM(tableIDColumnName_ )
		ENDIF 
		
		this.DebugMessage.Writeline( 'API.Append.Start', debug_ )
		this.DebugMessage.Writeline( selectCursor , debug_ )
		this.DebugMessage.Writeline( selectTable , debug_ )
		this.DebugMessage.Writeline( replaceID , debug_ )
		this.DebugMessage.Writeline( getIdValue , debug_ )
		this.DebugMessage.Writeline( useTableNameInZero , debug_ )

		
		* drop the curosr if it is in use.
		* not ... it would be unwise to use a cursor name the same as ann existing table.
		IF USED( cursorName_ ) then
			&selectCursor 
			use
		ENDIF 

		IF NOT USED( tableName_ ) then
			&useTableNameInZero 
			&selectTable 
		ENDIF 
					
		XMLTOCURSOR( xml_, cursorName_ )
		&selectCursor 
		*SELECT cursorName_ 
		
		rowIndex = 0 
		DO WHILE not EOF()
			this.DebugMessage.Writeline( 'appending row ' + STR(rowIndex) , debug_ )
			SCATTER MEMO MEMVAR 
			&selectTable
			*SELECT tableName_
			APPEND BLANK
			GATHER MEMO MEMVAR  
				
			this.DebugMessage.Writeline( 'gettting id ' , debug_ )
			&getIdValue
			this.DebugMessage.Writeline( 'id value is ' + STR( idValue ) , debug_ )
			replaceID = 'replace ' + ALLTRIM(cursorName_) + '.' + ALLTRIM(cursorIDColumnName_) + ' with ' + STR( idValue )
			this.DebugMessage.Writeline( replaceID, debug_ )
			&replaceID
			
			this.DebugMessage.Writeline( 'next row if any' , debug_ )
			&selectCursor
			*SELECT cursorName_
			SKIP 1
		ENDDO 
		commandText = 'select ' + tableName_
		
		CURSORTOXML( cursorName_ , "xmlText " , 1, 32, 0, '1')
		this.XML = xmlText 
		
		* note. only the last id is returned.
		* to do a bulk insert grab the id's from this.xml
		result = idValue
		
		this.DebugMessage.Writeline( 'API.Append.Finish', debug_ )
		RETURN result 
	ENDPROC

	PROCEDURE DebugMessageClear()	
		this.DebugMessage.Clear()
	ENDPROC
	
	PROCEDURE CarriageReturnLineFeed() as String 
		RETURN CHR(13) + CHR(10)
	ENDPROC
	
	PROCEDURE CreateIndexOnTable( tableName_ as String , expression_ as String, tagname_ as String ) 
		
		LOCAL selectTable as String 
		LOCAL tableIsInUse as Boolean
		LOCAL useTableWithExclusivePermissions as String 
		 
		TRY 
			selectTable = 'select ' + tableName_ 
			useTableWithExclusivePermissions = 'use ' + tableName_ + ' in 0 exclusive'
			IF USED( tableName_ ) THEN 
				tableIsInUse = .t.
				&selectTable
				use
			ELSE
				tableIsInUse = .f.
			ENDIF 
		CATCH TO exceptionInstance
		
		FINALLY
		
		ENDTRY 
		
	ENDPROC
	
	PROCEDURE ExecuteNonQueryDelete( sql_ as String ) as String 
		LOCAL result as String
		LOCAL commandText as String  
		TRY
			IF USED('deletedCountCursor') then
				SELECT deletedCountCursor
				use
			ENDIF 
			commandText = STRTRAN(LOWER(sql_), 'delete from', 'select count(*) as counted from')
			commandText = commandText + ' into cursor deletedCountCursor'
			&commandText 
			this._recordsAffected = deletedCountCursor.counted
			IF ( this._recordsAffected > 0 ) then
				&sql_
			ENDIF 
			USE IN deletedCountCursor
			result = ""
		CATCH TO exceptionInstance
			result = exceptionInstance.Message
			THROW result
		FINALLY  
		ENDTRY 
		RETURN result 
	ENDPROC 

	PROCEDURE ExecuteNonQueryInsert( sql_ as String ) as String 
		LOCAL result as String
		LOCAL commandText as String  
		TRY
			this._recordsAffected = 1 && how would we find this out ...
			IF ( this._recordsAffected > 0 ) then
				&sql_
			ENDIF 
			result = ""
		CATCH TO exceptionInstance
			result = exceptionInstance.Message
			THROW result 
		FINALLY  
		ENDTRY 
		RETURN result 
	ENDPROC 

	PROCEDURE ExecuteNonQueryUpdate( sql_ as String, tableName_ as string, where_ as String  ) as String 
		LOCAL result as String
		LOCAL commandText as String  
		TRY
			IF USED('updatedCountCursor') then
				SELECT updatedCountCursor
				use
			ENDIF 
			commandText = 'select count(*) as counted from ' + tableName_ + ' where ' + where_ + ' into cursor updatedCountCursor '
			&commandText 
			this._recordsAffected = updatedCountCursor.counted
			IF ( this._recordsAffected > 0 ) then
				&sql_
			ENDIF 
			USE IN updatedCountCursor
			result = ""
		CATCH TO exceptionInstance
			result = exceptionInstance.Message
			THROW result 
		FINALLY  
		ENDTRY 
		RETURN result 
	ENDPROC 

		
	PROCEDURE GetActualDefaultPath() as String 
		RETURN SYS(5)+SYS(2003)
	ENDPROC
	
	PROCEDURE GetXML( objectName_ as String, sql_ as String ) as string ;
		HELPSTRING "Returns the objects returned from the sql as xml"
		
		LOCAL result as String  
		LOCAL commandText as String 
		
		TRY 
			IF USED( objectName_ ) then
				commandText ='select ' + objectName_ 
				&commandText
				USE
			ENDIF 
		
			commandText = sql_ + ' into cursor ' + objectName_ 
			&commandText ;
			
			CURSORTOXML( objectName_ , "result" , 1, 32, 0, '1')
			
			commandText = 'use in ' + objectName_
			&commandText 
			
		CATCH TO exceptionInstance
			result = exceptioninstance.Message
		FINALLY
		
		ENDTRY 
		
		RETURN result 
		
	ENDPROC
	
	PROCEDURE RecordsAffected() as Integer 
		RETURN this._recordsAffected
	ENDPROC
	
	PROCEDURE SetDefaultTo( path_ as String )
		LOCAL commandText as string 
		commandText = 'set default to ' + path_
		&commandText
	ENDPROC
	
	PROCEDURE UpdateFromXML( tableName_ as string, tableIDColumnName_ as string, tagName_ as string, cursorName_ as string, cursorIDColumnName_ as string, xml_ as String, debug_ as Boolean  ) as bool  
		
		LOCAL result as Integer 
		LOCAL xmlText as String 
		LOCAL selectCursor as String 
		LOCAL selectTable as String 
		LOCAL useTableNameInZero as String 
		LOCAL replaceID as String 
		LOCAL rowIndex as Integer 
		LOCAL idValue as Integer 
		LOCAL getIdValue as String 
		LOCAL setOrderToTagName as String 

		
		this.SetDefaultTo( this.DefaultPath )
			
		selectCursor = 'select ' + cursorName_
		selectTable = 'select ' + tableName_ 
		replaceID = 'replace ' + ALLTRIM(cursorName_) + '.' + ALLTRIM(cursorIDColumnName_) + ' with ' + ALLTRIM(tableName_) + '.' + ALLTRIM(tableIDColumnName_ )
		useTableNameInZero = 'use ' + ALLTRIM(tableName_) + ' in 0'
		setOrderToTagName = 'set order to ' + tagName_ 
		
		getIdValue = 'idValue = ' + ALLTRIM(cursorName_) + '.' + ALLTRIM(cursorIDColumnName_) 

		this.debugMessage.Clear()
		
		this.DebugMessage.Writeline( 'UpdateFromXML.Start', debug_ )
		this.DebugMessage.Writeline( selectCursor , debug_ )
		this.DebugMessage.Writeline( selectTable , debug_ )
		this.DebugMessage.Writeline( replaceID , debug_ )
		this.DebugMessage.Writeline( getIdValue , debug_ )
		this.DebugMessage.Writeline( useTableNameInZero , debug_ )

		
		* drop the curosr if it is in use.
		* not ... it would be unwise to use a cursor name the same as ann existing table.
		IF USED( cursorName_ ) then
			&selectCursor 
			use
		ENDIF 

		IF NOT USED( tableName_ ) then
			&useTableNameInZero 
		ENDIF
		&selectTable  
		&setOrderToTagName 
		
		XMLTOCURSOR( xml_, cursorName_ )
		&selectCursor 
		rowIndex = 0 
		DO WHILE not EOF()
			this.DebugMessage.Writeline( 'updating row ' + STR(rowIndex) , debug_ )
			&getIdValue
		
			&selectTable
			SEEK idValue 
			
			IF FOUND() THEN 
				&selectCursor 
				SCATTER MEMO MEMVAR 
			
				&selectTable
				GATHER MEMO MEMVAR
				result = .t.
			ELSE 
				result = .f.
				THROW 'ID ' + STR( idValue ) + ' using tag ' + tagName_ + ' on table ' + tableName_ + ' was not found' 
			ENDIF 
			
			&selectCursor
			SKIP 1
		ENDDO 
		
		CURSORTOXML( cursorName_ , "xmlText " , 1, 32, 0, '1')
		this.XML = xmlText 
		
		this.DebugMessage.Writeline( 'UpdateFromXML.Finish', debug_ )
		RETURN result 
	ENDPROC
	
ENDDEFINE

DEFINE CLASS DebugMessage as Custom olepublic

	Message = ""
	
	PROCEDURE INIT
		
	ENDPROC 

	PROCEDURE Clear()
		this.message = ""
	ENDPROC 

	PROCEDURE Writeline( text_ as String, debug_ as Boolean )
		IF ( debug_ ) THEN 
			this.message = this.message + text_ + CHR(13) + CHR(10)
		ENDIF 
	ENDPROC
		
ENDDEFINE


DEFINE CLASS Queries as TableHelper olepublic

	PROCEDURE INIT
		SET RESOURCE off
		SET EXCLUSIVE off
		SET REPROCESS TO 2 seconds
		SET CPDIALOG off
		SET DELETED on
		SET EXACT off
		SET SAFETY off
		this.Name = 'adhoc_queries'
	ENDPROC 
	
	PROCEDURE CheckIndices() as Boolean
		LOCAL result as Boolean 
		result = .f.
		this.createindex( 'name', 'ADHOC_QUER', .f. )
		this.createindex( 'LEFT(name,40)', 'name', .f. )
		this.createindex( 'recno()', 'recno', .f. )
		this.CloseTable()
		result = .t.
		RETURN result			
	ENDPROC
	
ENDDEFINE 

DEFINE CLASS Suppliers as Session OLEPublic

	StartupPath= ""
	DefaultPath= "g:\tbsdata"
	
	PROCEDURE INIT
		SET RESOURCE off
		SET EXCLUSIVE off
		SET REPROCESS TO 2 seconds
		SET CPDIALOG off
		SET DELETED on
		SET EXACT off
		SET SAFETY off
		
		this.StartupPath= ADDBS(JUSTPATH(application.ServerName ))
		SET PATH TO ( this.StartupPath)
	ENDPROC
	
	PROCEDURE GetXML() as string ;
		HELPSTRING "Returns the suppliers as xml"
		
		LOCAL result as String  
		
		TRY 
			IF USED('supplierinstance') then
				SELECT supplierinstance
				USE
			ENDIF 
		
			SELECT * FROM supplier ORDER BY name INTO CURSOR supplierinstance
			CURSORTOXML( "supplierinstance", "result" , 1, 32, 0, '1')
			USE IN supplierinstance
		CATCH TO exceptionInstance
			result = exceptioninstance.Message
		FINALLY
		
		ENDTRY 
		
		RETURN result 
		
	ENDPROC
	
	PROCEDURE SetDefaultTo( path_ as String ) 
		LOCAL commandText as string 
		commandText = 'set default to ' + path_
		&commandText
	ENDPROC
	
ENDDEFINE


DEFINE CLASS SupplierColours as Session olepublic

	StartupPath= ""
	DefaultPath= "g:\tbsdata"
	
	PROCEDURE INIT
		SET RESOURCE off
		SET EXCLUSIVE off
		SET REPROCESS TO 2 seconds
		SET CPDIALOG off
		SET DELETED on
		SET EXACT off
		SET SAFETY off
		
		this.StartupPath= ADDBS(JUSTPATH(application.ServerName ))
		SET PATH TO ( this.StartupPath)
	ENDPROC
	
	PROCEDURE GetXML() as string ;
		HELPSTRING "Returns the suppliers colours as xml"
		
		LOCAL result as String  
		
		TRY
		
			IF USED('suppliercolourinstance') then
				SELECT suppliercolourinstance
				USE
			ENDIF 
		
			SELECT * FROM supplierColours ORDER BY id, code INTO CURSOR suppliercolourinstance
			CURSORTOXML( "suppliercolourinstance", "result" , 1, 32, 0, '1')
			USE IN suppliercolourinstance
		
		CATCH TO exceptionInstance
			result = exceptionInstance.Message	
		FINALLY
		
		ENDTRY
		
		RETURN result 
		
	ENDPROC
	
	PROCEDURE SetDefaultTo( path_ as String )
		*LPARAMETERS path_ as String 
		LOCAL commandText as string 
		*IF PARAMETERS() = 0 then
		*	path_ = _setDefaultTo
		*ELSE
		*	DefaultPath= path_ 		
		*ENDIF 
		commandText = 'set default to ' + path_
		&commandText
	ENDPROC
	
ENDDEFINE

DEFINE CLASS TableHelper as session olepublic

	CommandText = ""
	ExceptionMessage = "" 
	
	PROCEDURE INIT
		SET RESOURCE off
		SET EXCLUSIVE off
		SET REPROCESS TO 2 seconds
		SET CPDIALOG off
		SET DELETED on
		SET EXACT off
		SET SAFETY off
		IF ( not EMPTY(ALIAS())) THEN 
			this.Name = ALIAS()
		ENDIF 
	ENDPROC 

	PROCEDURE CheckIndices() as Boolean
		LOCAL result as Boolean 
		result = .f.
		
		RETURN result			
	ENDPROC
	
	PROCEDURE CloseTable()
		IF ( this.selecttable() ) then
			use
		ENDIF 
	ENDPROC
	
	PROCEDURE CreateIndex( expression_ as String , tagName_ as String, replace_ as Boolean ) as Boolean 
		LOCAL result as Boolean
		LOCAL commandText as String  
		result = .f.
		this.commandText = "" 
		TRY 
			IF ( this.selectTable() ) THEN 
				IF this.useExclusive() THEN 
					IF ( replace_ ) THEN 
						this.deleteIndex( tagName_ )
					ENDIF 
					IF NOT this.tagExists( tagName_ ) THEN
						commandText = 'index on ' + expression_ + ' tag ' + tagname_
						&commandText 
						result = .t.
					ENDIF 
				ENDIF 
			ENDIF
		CATCH TO exceptionInstance
			this.commandText = commandText
			this.ExceptionMessage = exceptionInstance.message
		FINALLY
		
		ENDTRY  
		RETURN result 
	ENDPROC
	
	
	PROCEDURE DeleteIndex( tagName_ as String ) as Boolean 
		LOCAL result as Boolean
		LOCAL commandText as String  
		result = .f.
		this.commandText = "" 
		TRY 
			IF this.tagExists( tagName_ ) THEN
				commandText = 'DELETE TAG ' + tagname_
				&commandText
			 	result = .t.
			ENDIF 
		CATCH TO exceptionInstance
			this.commandText = commandText
			this.exceptionMessage = exceptionInstance.message
		FINALLY
		
		ENDTRY  
		RETURN result 	
	ENDPROC
	
	PROCEDURE SelectTable() as Boolean 
		LOCAL result as Boolean
		LOCAL commandText as String  
		result = .t.
		this.exceptionMessage = ""
		TRY 
			IF NOT EMPTY( this.Name ) THEN 
				IF NOT USED( this.Name ) then
					commandText = 'USE ' + this.Name + ' IN 0'
					&commandText
				ENDIF 
				commandText = 'select ' + this.Name 
				&commandText 
				result = .t.
			ENDIF
		CATCH TO exceptionInstance
			this.commandText = commandText
			this.exceptionMessage = exceptionInstance.message
		FINALLY
		
		ENDTRY  
		RETURN result 
	ENDPROC
	
	PROCEDURE TagExists( tagName_ as String ) as Boolean 
		LOCAL result as Boolean 
		LOCAL tagIndex as Integer 
		result = .f.
		IF ( this.selectTable() ) THEN  
			FOR tagIndex = 1 TO TAGCOUNT()
				IF ( ALLTRIM( UPPER( tagName_) ) == ALLTRIM( UPPER( TAG( tagIndex ))) )THEN 
					result = .t.
					exit
				ENDIF 
			NEXT 
		ENDIF 
		RETURN result 
	ENDPROC
	
	PROCEDURE UseExclusive() as Boolean
		LOCAL result as Boolean 
		LOCAL commandText as String 
		result = .f.
		TRY
			IF NOT EMPTY( this.Name ) THEN 		
				IF USED( this.Name ) THEN 
					commandText = 'SELECT ' + this.Name
					&commandText
					use
				ENDIF 
				commandText = 'USE ' + this.Name + ' IN 0 EXCLUSIVE'
				&commandText 
				commandText = 'SELECT ' + this.Name
				&commandText 
				result = .t.
			ENDIF 
		CATCH TO exceptionInstance
			this.commandText = commandText
			this.exceptionMessage = exceptionInstance.message
		FINALLY
		
		ENDTRY
		RETURN result 
	ENDPROC 
	
	
ENDDEFINE  