IF OBJECT_ID('sp_CriaJOB') IS NOT NULL
   DROP PROCEDURE sp_CriaJOB

SET NOCOUNT ON
SET QUOTED_IDENTIFIER OFF

GO

CREATE PROCEDURE sp_CriaJOB 
(

   @JOB_Nome						VARCHAR(50),
   @JOB_Descricao					VARCHAR(255),   
   @JOB_Ativo						INT,
   @JOB_NotificaEmailQuando			INT,
   @JOB_NomePessoaNotificaEmail		VARCHAR(255),
   @JOB_EmailPessoaNotificaEmail	VARCHAR(255),
   @JOB_TipoComando					INT,
   @JOB_Comando						VARCHAR(8000),
   @JOB_DataInicio					VARCHAR(8),
   @JOB_HoraInicio					VARCHAR(6),
   @JOB_DataFim 					VARCHAR(8) = NULL,
   @JOB_HoraFim 					VARCHAR(6) = NULL,
   @JOB_FreqSubdayInterval		    INT,
   @JOB_FreqSubdayType			    INT,
   @JOB_FreqInterval				INT,   
   @LoginUsuarioServidor			VARCHAR(50),
   @Servidor						VARCHAR(50),
   @DataBase						VARCHAR(50)

) AS

BEGIN

   BEGIN TRANSACTION   

   SET NOCOUNT ON
   SET QUOTED_IDENTIFIER OFF

   DECLARE @NomeCategoria	  VARCHAR(255),
		   @NomeStep		  VARCHAR(255),
		   @NomeSchedule	  VARCHAR(255),
           @SubSystemComando  VARCHAR(10),
		   @JobID			  BINARY(16),
		   @ReturnCode		  INT 
      
   SELECT @ReturnCode    = 0,
          @NomeCategoria = '[JOB_CATEGORY_OF_CUSTOMIZACAO]',
		  @NomeStep      = 'JOB_STEP_' + @JOB_Nome,
		  @NomeSchedule  = 'JOB_SCHEDULE_' + @JOB_Nome

   IF @JOB_DataFim IS NULL 
	  SET @JOB_DataFim = '99991231'

   IF @JOB_HoraFim IS NULL 
	  SET @JOB_HoraFim = '235959'

   -- ******************************************************************************************************
   -- VERIFICAÇÃO DO PARÂMETRO "@JOB_TipoComando". SE O MESMO NÃO POSSUIR VALORES 0 OU 1, ENTÃO
   -- UMA MENSAGEM É EXIBIDA PARA USUÁRIO, INFORMANDO OS POSSIVEIS VALORES PARA O PARÂMETRO.
   -- EM SEGUIDA O FLUXO É DESVIADO PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
   -- SE O PARÂMETRO POSSUIR O VALOR ESPERADO, ENTÃO É ATRIBUIDO O VALOR CORRETO À VARIÁVEL 
   -- "@SubSystemComando".
   -- ******************************************************************************************************

   IF @JOB_Ativo NOT IN ( 0, 1 )
   BEGIN

	  RAISERROR( 'O PARÂMETRO "@JOB_Ativo" DEVE POSSUIR VALORES 1 (UM) OU 0(ZERO), ONDE:', 10, 1 )
	  RAISERROR( '', 10, 1 )
	  RAISERROR( '[ 0 = JOB INATIVO ];', 10, 1 )
	  RAISERROR( '[ 1 = JOB ATIVO ];', 10, 1 )
	  GOTO SAIDA_COM_ROLLBACK
	  
   END

   -- ******************************************************************************************************
   -- VERIFICAÇÃO DO PARÂMETRO "@JOB_TipoComando". SE O MESMO NÃO POSSUIR VALORES 0 OU 1, ENTÃO
   -- UMA MENSAGEM É EXIBIDA PARA USUÁRIO, INFORMANDO OS POSSIVEIS VALORES PARA O PARÂMETRO.
   -- EM SEGUIDA O FLUXO É DESVIADO PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
   -- SE O PARÂMETRO POSSUIR O VALOR ESPERADO, ENTÃO É ATRIBUIDO O VALOR CORRETO À VARIÁVEL 
   -- "@SubSystemComando".
   -- ******************************************************************************************************

   IF @JOB_TipoComando NOT IN ( 0, 1 )
   BEGIN

	  RAISERROR( 'O PARÂMETRO "@JOB_TipoComando" DEVE POSSUIR VALORES 1 (UM) OU 0(ZERO), ONDE:', 10, 1 )
	  RAISERROR( '', 10, 1 )
	  RAISERROR( '[ 0 = SCRIPT SQL ];', 10, 1 )
	  RAISERROR( '[ 1 = COMANDOS DO SISTEMA OPERACIONAL OU CHAMADA DE PROGRAMAS EXECUTÁVEIS ];', 10, 1 )
	  GOTO SAIDA_COM_ROLLBACK

   END ELSE 
   BEGIN 

	  IF @JOB_TipoComando = 0 
		 SET @SubSystemComando = 'TSQL'
	  ELSE
		 SET @SubSystemComando = 'CMDEXEC'
	  
   END

   -- ******************************************************************************************************
   -- VERIFICAÇÃO DA DATA DE INÍCIO E TERMINO DA EXECUÇÃO DO JOB.
   -- ******************************************************************************************************

   IF LTRIM( RTRIM( @JOB_DataFim ) ) <> ''
	  IF @JOB_DataInicio > @JOB_DataFim
	  BEGIN
		 RAISERROR( 'A DATA DE INÍCIO DA EXECUÇÃO DO JOB NÃO PODE SER MAIOR QUE A DATA FINAL DE EXECUÇÃO DO MESMO.', 10, 1 )
		 GOTO SAIDA_COM_ROLLBACK
	  END

   -- ******************************************************************************************************
   -- ABAIXO SEGUE A VERIFICAÇÃO DOS PARÂMETROS "@JOB_NomePessoaNotificaEmail" E "@JOB_NotificaEmailQuando".
   -- ******************************************************************************************************

   -- SE O NOME DE PESSOA PARA ENVIO DE EMAIL FOR INFORMADO NO PARÂMETRO "@JOB_NomePessoaNotificaEmail":
   IF LTRIM( RTRIM( @JOB_NomePessoaNotificaEmail ) ) <> ''
   BEGIN
	  
	  -- ENTÃO, VERIFICA SE O PARÂMETRO "@JOB_NotificaEmailQuando" NÃO ESTÁ ENTRE A FAIXA [1..3]:
	  IF @JOB_NotificaEmailQuando NOT IN ( 1, 2, 3 )
	  BEGIN

		 -- SENÃO ESTIVER ENTRE A FAIXA ACIMA, ENTÃO UMA MENSAGEM É EXIBIDA AO USUÁRIO
		 -- INFORMANDO OS VALORES CORRETOS A SEREM INFORMADOS. O FLUXO É DESVIADO PARA
		 -- O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
		 RAISERROR( 'QUANDO UMA PESSOA PARA ENVIO DE EMAIL É INFORMADA, O PARÂMETRO "@JOB_NotificaEmailQuando"', 10, 1 )
		 RAISERROR( 'DEVE RECEBER UM DOS SEGUINTES VALORES: ', 10, 1 )
		 RAISERROR( '', 10, 1 )
		 RAISERROR( '[ 1 = QUANDO O JOB TIVER SUCESSO ];', 10, 1 )
		 RAISERROR( '[ 2 = QUANDO O JOB FALHAR ];', 10, 1 )
		 RAISERROR( '[ 3 = QUANDO O JOB FOR CONCLUÍDO ].', 10, 1 )
		 RAISERROR( '', 10, 1 )
		 RAISERROR( 'O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )
		 GOTO SAIDA_COM_ROLLBACK 

	  END
      
   END ELSE
   BEGIN 

	  -- SE O NOME DE PESSOA PARA ENVIO DE EMAIL NÃO FOR INFORMADO NO PARÂMETRO "@JOB_NomePessoaNotificaEmail",
	  -- ENTÃO, É VERIFICADO SE O PARÂMETRO "@JOB_NotificaEmailQuando" POSSUI UM VALOR DIFERENTE DE 0 (ZERO).
	  IF @JOB_NotificaEmailQuando <> '0'
	  BEGIN

		 -- SE O PARÂMETRO "@JOB_NotificaEmailQuando" POSSUI UM VALOR DIFERENTE DE 0 (ZERO),
		 -- UMA MENSAGEM É EXIBIDA PARA USUÁRIO, INFORMANDO O VALOR CORRETO ESPERADO.
		 -- O FLUXO É DESVIADO PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
		 RAISERROR( 'QUANDO UMA PESSOA PARA ENVIO DE EMAIL NÃO É INFORMADA, O PARÂMETRO "@JOB_NotificaEmailQuando"', 10, 1 )
		 RAISERROR( 'DEVE RECEBER APENAS O VALOR [ 0 = NUNCA ]. O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )
		 GOTO SAIDA_COM_ROLLBACK

	  END

   END 

   -- ******************************************************************************************************
   -- ABAIXO SEGUE A VERIFICAÇÃO DO OPERADOR QUE PODERÁ RECEBER, OU NÃO, UM EMAIL DE CONFIRMAÇÃO DO JOB.
   -- ******************************************************************************************************

   -- SE O NOME DE PESSOA PARA ENVIO DE EMAIL FOR INFORMADO NO PARÂMETRO "@JOB_NomePessoaNotificaEmail":
   IF LTRIM( RTRIM( @JOB_NomePessoaNotificaEmail ) ) <> '' AND LTRIM( RTRIM( @JOB_EmailPessoaNotificaEmail ) ) <> ''
   BEGIN

	  -- VERIFICA SE EXISTE O OPERADOR INFORMADO NO PARÂMETRO "@JOB_NomePessoaNotificaEmail".   
	  IF EXISTS( SELECT * FROM msdb.dbo.sysoperators WHERE name = @JOB_NomePessoaNotificaEmail ) 
	  BEGIN

		 -- SE EXISTIR EXCLUI-SE ESSE OPERADOR.
		 EXECUTE @ReturnCode = msdb.dbo.sp_delete_operator @name = @JOB_NomePessoaNotificaEmail

		 -- SE A EXCLUSÃO FALHAR O FLUXO É DESVIADO PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
		 IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
		 BEGIN

			RAISERROR( 'O OPERADOR INFORMADO JÁ EXISTE E NÃO FOI POSSÍVEL EXCLUÍ-LO PARA SER RECRIADO', 10, 1 )
			RAISERROR( 'EM SEGUIDA, DEVIDO A UM ERRO DESCONHECIDO. O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )
			GOTO SAIDA_COM_ROLLBACK 

		 END ELSE 
		 BEGIN

			-- SE A EXCLUSÃO TIVER SUCESSO, ENTÃO O OPERADOR É ADICIONADO NOVAMENTE.
			EXECUTE @ReturnCode = msdb.dbo.sp_add_operator @name          = @JOB_NomePessoaNotificaEmail,
														   @email_address = @JOB_EmailPessoaNotificaEmail,
	  													   @enabled       = 1
   	  
			-- SE A ADIÇÃO FALHAR O FLUXO É DESVIADO PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
			IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
			BEGIN

			   RAISERROR( 'NÃO FOI POSSÍVEL CRIAR O OPERADOR INFORMADO NO PARÂMETRO ''@JOB_NomePessoaNotificaEmail''. O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )
			   GOTO SAIDA_COM_ROLLBACK 

			END

		 END

	  END ELSE
	  BEGIN
   	  
		 -- SE O OPERADOR NÃO EXISTIR, ENTÃO O MESMO É ADICIONADO.
		 EXECUTE @ReturnCode = msdb.dbo.sp_add_operator @name          = @JOB_NomePessoaNotificaEmail,
													    @email_address = @JOB_EmailPessoaNotificaEmail,
	  												    @enabled       = 1

		 -- SE A ADIÇÃO FALHAR O FLUXO É DESVIADO PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
		 IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
		 BEGIN

			RAISERROR( 'NÃO FOI POSSÍVEL CRIAR O OPERADOR INFORMADO NO PARÂMETRO ''@JOB_NomePessoaNotificaEmail''. O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )
			GOTO SAIDA_COM_ROLLBACK	
   	  
		 END

	  END

   END

   -- ******************************************************************************************************
   -- ABAIXO SEGUE A VERIFICAÇÃO DA CATEGORIA DE JOB DESTINADA AOS JOBS DA CUSTOMIZAÇÃO. 
   -- CASO NÃO EXISTA ESSA CATEGORIA É ADICIONADA.
   -- ******************************************************************************************************

   -- VERIFICA SE JÁ EXISTE UMA CATEGORIA DE JOB COM O NOME ESPECIFICADO NA 
   -- VARIÁVEL "@NomeCategoria" SE NÃO EXISTIR, ENTÃO CRIA A CATEGORIA PARA O JOB A SER CRIADO.
   IF ( SELECT COUNT( * ) FROM msdb.dbo.syscategories WHERE name = @NomeCategoria ) < 1 
   BEGIN

	  EXECUTE @ReturnCode = msdb.dbo.sp_add_category @class = 'JOB',
													 @type  = 'LOCAL',
													 @name  = @NomeCategoria
	  	  
	  -- SE A ADIÇÃO FALHAR O FLUXO É DESVIADO PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
	  IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
	  BEGIN

		 EXECUTE( "RAISERROR( 'NÃO FOI POSSÍVEL CRIAR A CATEGORIA DE JOBS CUSTOMIZADOS CHAMADA: ''" + @NomeCategoria + "''. O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )" )
		 GOTO SAIDA_COM_ROLLBACK	
	  
	  END

   END

   -- ******************************************************************************************************
   -- VERIFICAÇÃO DO JOB A SER CRIADO. CASO EXISTA, É VERIFICADO O TIPO DESTE JOB. 
   -- AO FINAL, SE O JOB EXISTIR E NÃO FOR DO TIPO MULTI-SERVER, ENTÃO ELE É DELETADO.
   -- ******************************************************************************************************

   -- VERIFICA SE JÁ EXISTE O JOB A SER CRIADO.
   SELECT @JobID = job_id FROM msdb.dbo.sysjobs 
   WHERE name = @JOB_Nome

   -- SE O JOB JÁ EXISTIR.
   IF @JobID IS NOT NULL
   BEGIN 
	  
	  -- CHECA SE O JOB É UM JOB DO TIPO MULTI-SERVER 
	  IF EXISTS( SELECT * FROM msdb.dbo.sysjobservers WHERE job_id = @JobID AND server_id <> 0 ) 
	  BEGIN

		 -- SE O JOB FOR MULTI-SERVER EXIBE UMA MENSAGEM DE ERRO E ABORTA
		 -- A TRANSAÇÃO, DESVIANDO PARA O LABEL "SAIDA_COM_ROLLBACK".
		 EXECUTE( "RAISERROR ( 'O JOB ''" + @JOB_Nome + "'' JÁ EXISTE, SENDO IMPOSSÍVEL RECRIÁ-LO, POIS ESTE JOB É DO TIPO MULTI-SERVER. TENTE CRIAR O JOB COM UM NOME DIFERENTE.', 10, 1 )" )
		 GOTO SAIDA_COM_ROLLBACK  

	  END ELSE 
	  BEGIN

		 -- SE O JOB NÃO FOR DO TIPO MULTI-SERVER, ENTÃO ELE É DELETADO.
		 EXECUTE @ReturnCode = msdb.dbo.sp_delete_job @job_name               = @JOB_Nome,
                                                      @delete_unused_schedule = 1
		 
		 -- SE A DELEÇÃO DO JOB FALHAR O FLUXO É DESVIADO PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
		 IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
		 BEGIN

			EXECUTE( "RAISERROR( 'NÃO FOI POSSÍVEL CRIAR A CATEGORIA DE JOBS CUSTOMIZADOS CHAMADA: ''" + @NomeCategoria + "''. O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )" )
			GOTO SAIDA_COM_ROLLBACK	

		 END

	  END

	  -- SETA COMO NULA A VARIÁVEL UTILIZADA PARA GUARDAR O ID DO JOB VERIFICADO ACIMA.
	  SET @JobID = NULL

   END
   
   -- ******************************************************************************************************
   -- ABAIXO O JOB É ADICIONADO COM OS PARÂMETROS INFORMADOS.
   -- ******************************************************************************************************

   -- REALIZA TODOS OS PASSOS PARA A CRIÇÃO DO JOB, ADICIONANDO O MESMO 
   -- AO SERVIDOR, ADICIONANDO OS SEUS PASSOS E TAREFAS A SEREM REALIZADAS.
   BEGIN 

	  -- TENTA ADICIONAR O CABEÇALHO DO JOB.
	  EXECUTE @ReturnCode = msdb.dbo.sp_add_job @job_id                     = @JobID OUTPUT, 
												@job_name                   = @JOB_Nome, 
												@owner_login_name           = @LoginUsuarioServidor, 
												@description                = @JOB_Descricao, 
												@category_name              = @NomeCategoria, 
												@enabled                    = @JOB_Ativo, 
												@notify_level_email         = @JOB_NotificaEmailQuando, 
												@notify_email_operator_name = @JOB_NomePessoaNotificaEmail,
												@notify_level_eventlog      = 2, 
												@delete_level               = 0
	  
	  -- CASO OCORRA UM ERRO, UMA MENSAGEM É EXIBIDA AO USUÁRIO E O FLUXO É DESVIADO PARA O 
	  -- LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA. 
	  IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
	  BEGIN
		 
		 RAISERROR( 'NÃO FOI POSSÍVEL CRIAR O CABEÇALHO DO JOB. O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )
		 GOTO SAIDA_COM_ROLLBACK 
	  
	  END ELSE 
	  BEGIN

		 -- SE O CABEÇALHO DO JOB FOI ADICIONADO, ENTÃO TENTA ADICIONAR O JOB STEP. 
		 EXECUTE @ReturnCode = msdb.dbo.sp_add_jobstep @job_id               = @JobID,  
													   @step_id              = 1, 
													   @step_name            = @NomeStep, 
													   @command              = @JOB_Comando, 
													   @database_name        = @DataBase, 
													   @subsystem            = @SubSystemComando

		 -- CASO OCORRA UM ERRO, UMA MENSAGEM É EXIBIDA AO USUÁRIO E O FLUXO É DESVIADO PARA O 
		 -- LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA. 
		 IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
		 BEGIN
   		 
			RAISERROR( 'NÃO FOI POSSÍVEL CRIAR O CABEÇALHO DO JOB. O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )
			GOTO SAIDA_COM_ROLLBACK 
   	  
		 END ELSE 
		 BEGIN

			-- TENTA ATUALIZAR O JOB.
			EXECUTE @ReturnCode = msdb.dbo.sp_update_job @job_id        = @JobID, 
													     @start_step_id = 1 

			-- CASO OCORRA UM ERRO, UMA MENSAGEM É EXIBIDA AO USUÁRIO E O FLUXO É DESVIADO PARA O 
			-- LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA. 
			IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
			BEGIN
      		 
			   RAISERROR( 'NÃO FOI POSSÍVEL CRIAR O CABEÇALHO DO JOB. O PROCESSO SERÁ INTERROMPIDO.', 10, 1 )
			   GOTO SAIDA_COM_ROLLBACK 
      	  
			END ELSE 
			BEGIN

			   -- TENTA ADICIONAR O JOB SCHEDULE, CASO OCORRA UM ERRO, O FLUXO É 
			   -- DESVIADO PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
			   EXECUTE @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id                 = @JobID, 
															     @name                   = @NomeSchedule, 
															     @enabled                = 1, 
															     @freq_type              = 8, 
															     @active_start_date      = @JOB_DataInicio, 
															     @active_start_time      = @JOB_HoraInicio,
																 @active_end_date        = @JOB_DataFim,
																 @active_end_time        = @JOB_HoraFim,
															     @freq_interval          = @JOB_FreqInterval, 
															     @freq_subday_type       = @JOB_FreqSubdayType, 
															     @freq_subday_interval   = @JOB_FreqSubdayInterval, 
															     @freq_relative_interval = 0,
															     @freq_recurrence_factor = 1 

			   -- CASO OCORRA UM ERRO, UMA MENSAGEM É EXIBIDA AO USUÁRIO E O FLUXO É DESVIADO PARA O 
			   -- LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA. 
			   IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
			   BEGIN
         		 
				  RAISERROR( 'NÃO FOI POSSÍVEL CRIAR O AGENDAMENTO DO JOB.', 10, 1 )
				  GOTO SAIDA_COM_ROLLBACK 
         	  
			   END ELSE 
               BEGIN

				  -- TENTA ADICIONAR O JOB SERVER, CASO OCORRA UM ERRO, O FLUXO É DESVIADO
				  -- PARA O LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA.
				  EXECUTE @ReturnCode = msdb.dbo.sp_add_jobserver @job_id      = @JobID, 
															      @server_name = @Servidor	

				  -- CASO OCORRA UM ERRO, UMA MENSAGEM É EXIBIDA AO USUÁRIO E O FLUXO É DESVIADO PARA O 
				  -- LABEL "SAIDA_COM_ROLLBACK" E A TRANSAÇÃO É CANCELADA. 
				  IF ( @@ERROR <> 0 OR @ReturnCode <> 0 ) 
				  BEGIN

					 RAISERROR( 'NÃO FOI POSSÍVEL ADICIONAR O SERVIDOR PARA EXECUÇÃO DO JOB.', 10, 1 )
					 GOTO SAIDA_COM_ROLLBACK 

				  END

			   END

			END

		 END 

	  END

   END

   COMMIT TRANSACTION
   GOTO SAIDA_COM_SUCESSO
           
   SAIDA_COM_ROLLBACK: IF ( @@TRANCOUNT > 0 ) ROLLBACK TRANSACTION 
   SAIDA_COM_SUCESSO:

END 

/*

EXECUTE sp_CriaJOB @JOB_Nome                     = 'JOB_INTEGRADOR_SUBIDA', 
				   @JOB_Descricao                = 'Atualização do e-Commerce Ikeda', 
				   @JOB_Ativo                    = 1, 
				   @JOB_NotificaEmailQuando      = 0,
				   @JOB_NomePessoaNotificaEmail  = '',
				   @JOB_EmailPessoaNotificaEmail = '',
				   @JOB_TipoComando              = 1,
				   @JOB_Comando                  = '"c:\windows\system\notepad.exe"',
				   @JOB_DataInicio               = '20110927',
				   @JOB_HoraInicio               = '60000',
				   @JOB_DataFim                  = null,
				   @JOB_HoraFim                  = '220000',
				   @JOB_FreqInterval             = 127,
				   @JOB_FreqSubdayInterval       = 8,
                   @JOB_FreqSubdayType           = 8,                   
				   @LoginUsuarioServidor         = 'sa',
				   @DataBase                     = 'DBBIBELOT',
				   @Servidor                     = '(local)'

|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_Nome                     |   NOME DO JOB;                                                                                       |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_Descricao                |   DESCRIÇÃO DO JOB;                                                                                  |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_Ativo                    |   0 - ATIVO;                                                                                         |
|                              |   1 - DESATIVO.                                                                                      |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_NotificaEmailQuando      |   0 - NUNCA ( O PARÂMETRO "@JOB_NomePessoaNotificaEmail" DEVE SER EM BRANCO );                       |
|                              |   1 - QUANDO O JOB TIVER SUCESSO ( O PARÂMETRO "@JOB_NomePessoaNotificaEmail" DEVE SER PREENCHIDO ); |
|                              |   2 - QUANDO O JOB FALHAR ( O PARÂMETRO "@JOB_NomePessoaNotificaEmail" DEVE SER PREENCHIDO );        |
|                              |   3 - QUANDO O JOB FOR CONCLUÍDO ( O PARÂMETRO "@JOB_NomePessoaNotificaEmail" DEVE SER PREENCHIDO ). |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_NomePessoaNotificaEmail  |   NOME DA PESSOA QUE RECEBERÁ O EMAIL CONFORNE O PARÂMETRO "@JOB_NotificaEmailQuando".               |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@job_EmailPessoaNotificaEmail |   EMAIL DA PESSOA QUE RECEBERÁ O EMAIL CONFORNE O PARÂMETRO "@JOB_NotificaEmailQuando".              |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_TipoComando              |   0 - SCRIPT SQL;                                                                                    |
|                              |   1 - COMANDOS DO SISTEMA OPERACIONAL OU CHAMADA DE PROGRAMAS EXECUTÁVEIS.                           |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_Comando                  |   COMANDO SQL OU DE SISTEMA OPERACIONAL OU O JOB REALIZARÁ.                                          |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_DataInicio               |   DATA DE INÍCIO DA EXECUÇÃO DO JOB (FORMATO: YYYYMMDD).                                             |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_HoraInicio               |   HORA DE INÍCIO DA EXECUÇÃO DO JOB (FORMATO: HHMMSS).                                               |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_DataFim                  |   DATA DE TÉRMINO DA EXECUÇÃO DO JOB (FORMATO: YYYYMMDD).                                            |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_HoraFim                  |   HORA DE TÉRMINO DA EXECUÇÃO DO JOB (FORMATO: HHMMSS, SUPRIMIDO ZEROS à ESQUERDA).                  |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_FreqInterval             |   VALOR INTEIRO QUE REPRESENTA OS DIAS DA SEMANA EM QUE O JOB SERÁ EXECUTADO, SENDO                  |
|                              |   DOMINGO = 1                                                                                        |
|                              |   SEGUNDA = 2                                                                                        |
|                              |   TERÇA   = 4                                                                                        |
|                              |   QUARTA  = 8                                                                                        |
|                              |   QUINTA  = 16                                                                                       |
|                              |   SEXTA   = 32                                                                                       |
|                              |   SÁBADO  = 64; PODENDO SER UTILIZADO QUALQUER UM DESSES VALORES COMBINADOS COM O OPERADOR "OR",     |
|                              |                 EXEMPLO: EXECUÇÃO NO DOMINGO E NA QUINTA -> DOMINGO "OR" QUINTA -> 1 OR 16 -> 17.    |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_FreqSubdayInterval       |   QUANTIDADE DE UNIDADES (VER PARÂMETRO "@JOB_FreqSubdayType") QUE O JOB VAI SER EXECUTADO NOS       |
|                              |   DIAS SELECIONADOS NO PARÂMETRO "@JOB_FreqInterval", SENDO:                                         |
|                              |   [1..60] PARA O PARÂMETRO "@JOB_FreqSubdayType" = [ 4 - MINUTOS ];                                  |
|                              |   [1..24] PARA O PARÂMETRO "@JOB_FreqSubdayType" = [ 8 - HORAS ]                                     |
|-------------------------------------------------------------------------------------------------------------------------------------|
|@JOB_FreqSubdayType           |   UNIDADE DE MEDIDA PARA QUANTIDADE A SER EXECUTADA, SENDO:                                          |
|                              |   MINUTOS  = 4;                                                                                      |
|                              |   HORAS    = 8.                                                                                      |
|-------------------------------------------------------------------------------------------------------------------------------------|

*/