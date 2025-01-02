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
                                      @Valor_Transacao = 170.85, 
                                      @Descricao = 'Compra no ARAGUAIA shopping';

        -- Excluir uma transação (soft delete):
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'DELETE', @Id_Transacao = 1;

        -- Consultar transações:
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Id_Cliente = 1';
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Numero_Cartao ='1111111111111111'
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Data_Inicial = '2024-01-01', @Data_Final = '2024-12-31';
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Valor_Transacao = 100.00;
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Descricao = '%SUPERMERCADO%';
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Numero_Cartao ='2222222222222222', @Descricao = '%SUPERMERCADO%';
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Numero_Cartao ='1111111111111111', @Data_Inicial = '2024-01-01', @Data_Final = '2025-12-31';
 ====================================================================================================================================================== */
CREATE PROCEDURE sp_CartaoTransacoes_CRUD
    @Acao VARCHAR(7), -- Ações (INSERT, UPDATE, DELETE, SELECT)
    @Id_Transacao INT = NULL, 
    @Id_Cliente INT = NULL,
    @Numero_Cartao VARCHAR(16) = NULL,
    @Valor_Transacao DECIMAL(12, 2) = NULL,
    @Data_Inicial DATETIME = NULL, -- Parâmetro de data inicial
    @Data_Final DATETIME = NULL,   -- Parâmetro de data final
    @Descricao VARCHAR(100) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Erro INT = 0;

    BEGIN TRY
        -- Inicia transação no banco de dados
        BEGIN TRANSACTION;

        -- Insere nova transação de cartão (Ação: INSERT)
		-- Para esta ação cpnsiderar o @Data_Inicial como Data_Transação
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

        -- Altera transação de cartão (Ação: UPDATE)
		-- Para esta ação considerar o paramêtro @Data_Inicial como Data_Transação
        IF @Acao = 'UPDATE' 
        BEGIN
            IF @Id_Transacao IS NULL
            BEGIN
                SET @Erro = 1;
                THROW 50000, 'Id_Transacao inválido, campo obrigatório para atualização.', 1;
            END;

			-- Antes de agir verifica se a transação existe
			IF NOT EXISTS (SELECT 1 FROM Cartao_Transacoes WHERE Id_Transacao = @Id_Transacao) 
            BEGIN
                THROW 50000, 'Transação não encontrada para atualização.', 1;
            END;

            UPDATE Cartao_Transacoes
            SET
                Numero_Cartao = COALESCE(@Numero_Cartao, Numero_Cartao),
                Valor_Transacao = COALESCE(@Valor_Transacao, Valor_Transacao),
                Data_Transacao = COALESCE(@Data_Inicial, Data_Transacao),
                Descricao = COALESCE(@Descricao, Descricao)
            WHERE Id_Transacao = @Id_Transacao;
        END;

        -- Soft Delete da transação de cartão - Marca a transação como inativa 'I' (Ação: DELETE)
        IF @Acao = 'DELETE' 
        BEGIN 
            IF @Id_Transacao IS NULL
            BEGIN
                SET @Erro = 1;
                THROW 50000, 'Id_Transacao inválido, campo obrigatório para exclusão.', 1;
            END;

			-- Antes de agir verifica se a transação existe
            IF NOT EXISTS (SELECT 1 FROM Cartao_Transacoes WHERE Id_Transacao = @Id_Transacao)
            BEGIN
                THROW 50000, 'Transação não encontrada para exclusão.', 1;
            END;

            UPDATE Cartao_Transacoes
            SET Status = 'I'
            WHERE Id_Transacao = @Id_Transacao;
        END;

        -- Consulta transações por diversos parâmetros (Ação: SELECT)
        IF @Acao = 'SELECT' 
        BEGIN
            SELECT Id_Transacao, Id_Cliente, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status
            FROM Cartao_Transacoes
            WHERE (@Id_Cliente IS NULL OR Id_Cliente = @Id_Cliente)
              AND (@Numero_Cartao IS NULL OR Numero_Cartao = @Numero_Cartao)
              AND (@Valor_Transacao IS NULL OR Valor_Transacao = @Valor_Transacao)
              AND (@Data_Inicial IS NULL AND @Data_Final IS NULL OR Data_Transacao BETWEEN @Data_Inicial AND @Data_Final) -- Período entre duas datas
              AND Status = 'A' -- Só transações ativas
        END;

        -- Se tudo correr bem, comita a transação no banco de dados
        COMMIT TRANSACTION;
    END TRY

    BEGIN CATCH
        -- Em caso de erro, realiza o rollback da transação no banco de dados
        ROLLBACK TRANSACTION;
        SELECT ERROR_MESSAGE() AS ErrorMessage;
    END CATCH
END;
