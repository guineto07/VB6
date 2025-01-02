/*================================================================================================================================================================
    AUTOR: Guilherme Neto
    DATA DE CRIAÇÃO: 31/12/2024
    DESCRIÇÃO: 
        Esta procedure permite realizar operações CRUD na tabela Cartao_Transacoes, a exclusão é feita de forma lógica (soft delete),
        consultas de transações com base nos parâmetros Id_Cliente, Numero_Cartao, intevalo de Data_Transacao e pelo Valor_Transacao.
		Para as ações INSERT, UPDATE e DELETE, é necessário que o cliente e o cartão correspondente já esteja cadastrado na tabela Clientes.
		
    PARÂMETROS:
        @Acao VARCHAR(7)          - Especifica a ação a ser realizada: 'INSERT', 'UPDATE', 'DELETE', 'SELECT'.
        @Id_Transacao INT         - Identificador da transação a ser modificada (necessário para UPDATE e DELETE).
        @Id_Cliente INT           - Identificador do cliente (utilizado nas ações INSERT, UPDATE, DELETE e SELECT). Chave estrangeira com a tabela Clientes.
        @Numero_Cartao VARCHAR(16) - Número do cartão associado à transação.
        @Valor_Transacao DECIMAL(12, 2) - Valor da transação.
        @Descricao VARCHAR(100)   - Descrição detalhada da transação.
        @Data_Inicial DATETIME    - Data inicial para filtro de transações (usado na consulta SELECT) e considerado como Data_Transacao nas ações INSERT e UPDATE.
        @Data_Final DATETIME      - Data final para filtro de transações (usado na consulta SELECT).

    ERROS POSSÍVEIS:
        - Caso @Id_Transacao não seja fornecido para operações de atualização ou exclusão, um erro será gerado.
        - Caso o @Id_Cliente não exista na tabela Clientes, um erro será gerado nas operações de 'INSERT', 'UPDATE' ou 'DELETE'.
        - Se ocorrer uma falha durante a execução de qualquer operação, a transação será revertida (rollback).

    EXEMPLOS DE USO:
        -- Inserir uma nova transação:
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'INSERT', @Id_Cliente = 1, 
                                      @Numero_Cartao = '1111111111111111', 
                                      @Valor_Transacao = 100.00, 
                                      @Data_Inicial = '2024-12-31', 
                                      @Descricao = 'Compra no supermercado';

        -- Alterar uma transação existente:
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'UPDATE', @Id_Transacao = 1, 
                                      @Valor_Transacao = 10900.99, 
                                      @Descricao = 'Viagem USA shopping';

        -- Excluir uma transação (soft delete):
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'DELETE', @Id_Transacao = 1;

        -- Consultar transações:

        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Id_Transacao= 1;
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Id_Cliente = 1;
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Numero_Cartao ='1111111111111111'
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Data_Inicial = '2024-01-01', @Data_Final = '2024-01-31';
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Valor_Transacao = 170.85;
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Descricao = 'Compra%';
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Numero_Cartao ='2222222222222222', @Descricao = '%SUPERMERCADO%';
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Numero_Cartao ='1111111111111111', @Data_Inicial = '2024-12-01', @Data_Final = '2024-12-31';
 ====================================================================================================================================================== */
CREATE PROCEDURE sp_CartaoTransacoes_CRUD
    @Acao VARCHAR(7),  -- Ações (INSERT, UPDATE, DELETE, SELECT)
    @Id_Transacao INT = NULL, 
    @Id_Cliente INT = NULL,
    @Numero_Cartao VARCHAR(16) = NULL,
    @Valor_Transacao DECIMAL(12, 2) = NULL,
    @Data_Inicial DATETIME = NULL,  -- Parâmetro de data inicial
    @Data_Final DATETIME = NULL,    -- Parâmetro de data final
    @Descricao VARCHAR(100) = NULL,
    @Status_Sp INT = 0 OUTPUT,             -- Parâmetro de saída (Status 0 = sucesso, 1 = erro)
    @ErrorMessage VARCHAR(500) ='' OUTPUT,  -- Parâmetro de saída (Mensagem de erro)
    @ErrorNumber INT = 0 OUTPUT              -- Parâmetro de saída (Número do erro)
AS
BEGIN
    SET NOCOUNT ON;

    -- Inicializar parâmetros de saída
    SET @Status_Sp = 0;
    SET @ErrorMessage = '';
    SET @ErrorNumber = 0;

    -- Verificar se a ação foi informada
    IF @Acao IS NULL OR LTRIM(RTRIM(@Acao)) = ''
    BEGIN
        SET @Status_Sp = 1;
        SET @ErrorMessage = 'Parâmetro @Acao é obrigatório.';
        SET @ErrorNumber = 50001; -- Código de erro customizado
        THROW 50000, @ErrorMessage, 1;
        RETURN;
    END

	-- Caso o parâmetro de ação seja 'SELECT', não é necessário Id_Cliente ou Número_Cartão, mas ainda podemos verificar se necessário
    IF  (@Acao = 'INSERT' OR @Acao = 'SELECT') 
		AND (@Id_Transacao IS NULL AND @Id_Cliente IS NULL AND @Numero_Cartao IS NULL 
	    AND  @Data_Inicial IS NULL AND @Data_Final IS NULL AND @Valor_Transacao IS NULL AND @Descricao IS NULL)
    BEGIN
        SET @Status_Sp = 1;
        SET @ErrorMessage = 'Pelo menos um dos parâmetros @Id_Transacao, @Id_Cliente, @Valor_transacao,
							 @Numero_Cartao, @Descricao, @Valor, @Data_Inicial e @Data_Final deve ser informado para a consulta.';
        SET @ErrorNumber = 50003; -- Código de erro customizado
        THROW 50000, @ErrorMessage, 1;
        RETURN;
    END

    -- Verificar se os parâmetros obrigatórios estão informados dependendo da ação
	IF @Acao IN ('UPDATE', 'DELETE') AND (@Id_Transacao IS NULL)
    BEGIN
        SET @Status_Sp = 1;
        SET @ErrorMessage = 'Parâmetro @Id_Transacao obrigatório para UPDATE e DELETE.';
        SET @ErrorNumber = 50002; -- Código de erro customizado
        THROW 50000, @ErrorMessage, 1;
        RETURN;
    END

    BEGIN TRY
        -- Inicia transação no banco de dados
        BEGIN TRANSACTION;

        -- Ações de inserção, atualização, exclusão, etc.

        -- Inserção de transação 
        IF @Acao = 'INSERT' 
        BEGIN
            -- Antes de agir verifica se a transação existe
            IF NOT EXISTS (SELECT 1 FROM Clientes WHERE Numero_Cartao = @Numero_Cartao)
            BEGIN
                THROW 50000, 'Número de cartão não encontrado.', 1;
            END;

            INSERT INTO Cartao_Transacoes (Id_Cliente, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao)
            VALUES (@Id_Cliente, @Numero_Cartao, @Valor_Transacao, @Data_Inicial, @Descricao);
        END;

        -- Atualização de transação 
        IF @Acao = 'UPDATE' 
        BEGIN
            IF @Id_Transacao IS NULL
            BEGIN
                SET @Status_Sp = 1;
                SET @ErrorMessage = 'Id_Transacao inválido, campo obrigatório para atualização.';
                SET @ErrorNumber = 50004; -- Código de erro customizado
                THROW 50000, @ErrorMessage, 1;
            END;

            -- Verifica se a transação existe
            IF NOT EXISTS (SELECT 1 FROM Cartao_Transacoes WHERE Id_Transacao = @Id_Transacao)
            BEGIN
                SET @Status_Sp = 1;
                SET @ErrorMessage = 'Transação não encontrada para atualização.';
                SET @ErrorNumber = 50005; -- Código de erro customizado
                THROW 50000, @ErrorMessage, 1;
            END;

            UPDATE Cartao_Transacoes
            SET
                Numero_Cartao = COALESCE(@Numero_Cartao, Numero_Cartao),
                Valor_Transacao = COALESCE(@Valor_Transacao, Valor_Transacao),
                Data_Transacao = COALESCE(@Data_Inicial, Data_Transacao),
                Descricao = COALESCE(@Descricao, Descricao)
            WHERE Id_Transacao = @Id_Transacao;
        END;

        -- Soft Delete da transação 
        IF @Acao = 'DELETE' 
        BEGIN 
            IF @Id_Transacao IS NULL
            BEGIN
                SET @Status_Sp = 1;
                SET @ErrorMessage = 'Id_Transacao inválido, campo obrigatório para exclusão.';
                SET @ErrorNumber = 50006; -- Código de erro customizado
                THROW 50000, @ErrorMessage, 1;
            END;

            -- Verifica se a transação existe
            IF NOT EXISTS (SELECT 1 FROM Cartao_Transacoes WHERE Id_Transacao = @Id_Transacao)
            BEGIN
                SET @Status_Sp = 1;
                SET @ErrorMessage = 'Transação não encontrada para exclusão.';
                SET @ErrorNumber = 50007; -- Código de erro customizado
                THROW 50000, @ErrorMessage, 1;
            END;

            UPDATE Cartao_Transacoes
            SET Status = 'I'
            WHERE Id_Transacao = @Id_Transacao;
        END;

        -- Consulta transações (Ação: SELECT)
        IF @Acao = 'SELECT' 
        BEGIN
            SELECT Id_Transacao, Id_Cliente, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status
            FROM Cartao_Transacoes
            WHERE (@Id_Transacao IS NULL OR Id_Transacao = @Id_Transacao)
		    	AND (@Id_Cliente IS NULL OR Id_Cliente = @Id_Cliente)
              AND (@Numero_Cartao IS NULL OR Numero_Cartao = @Numero_Cartao)
              AND (@Valor_Transacao IS NULL OR Valor_Transacao = @Valor_Transacao)
              AND (@Data_Inicial IS NULL AND @Data_Final IS NULL OR Data_Transacao BETWEEN @Data_Inicial AND @Data_Final) -- Período entre duas datas
              AND (@Descricao IS NULL OR Descricao LIKE @Descricao)
			 AND Status = 'A' -- Só transações ativas
        END;

        -- Se tudo correr bem, comita a transação no banco de dados
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        -- Em caso de erro, faz o rollback e retorna a mensagem de erro
        ROLLBACK TRANSACTION;
        SET @Status_Sp = 1;
        SET @ErrorMessage = ERROR_MESSAGE();
        SET @ErrorNumber = ERROR_NUMBER();
    END CATCH
END;

