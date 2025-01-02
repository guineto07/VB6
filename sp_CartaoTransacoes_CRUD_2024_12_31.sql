/*================================================================================================================================================================
    AUTOR: Guilherme Neto
    DATA DE CRIA��O: 31/12/2024
    DESCRI��O: 
        Esta procedure permite realizar opera��es CRUD na tabela Cartao_Transacoes, a exclus�o � feita de forma l�gica (soft delete),
        consultas de transa��es com base nos par�metros Id_Cliente, Numero_Cartao, intevalo de Data_Transacao e pelo Valor_Transacao.
		Para as a��es INSERT, UPDATE e DELETE, � necess�rio que o cliente e o cart�o correspondente j� esteja cadastrado na tabela Clientes.
		
    PAR�METROS:
        @Acao VARCHAR(7)          - Especifica a a��o a ser realizada: 'INSERT', 'UPDATE', 'DELETE', 'SELECT'.
        @Id_Transacao INT         - Identificador da transa��o a ser modificada (necess�rio para UPDATE e DELETE).
        @Id_Cliente INT           - Identificador do cliente (utilizado nas a��es INSERT, UPDATE, DELETE e SELECT). Chave estrangeira com a tabela Clientes.
        @Numero_Cartao VARCHAR(16) - N�mero do cart�o associado � transa��o.
        @Valor_Transacao DECIMAL(12, 2) - Valor da transa��o.
        @Descricao VARCHAR(100)   - Descri��o detalhada da transa��o.
        @Data_Inicial DATETIME    - Data inicial para filtro de transa��es (usado na consulta SELECT) e considerado como Data_Transacao nas a��es INSERT e UPDATE.
        @Data_Final DATETIME      - Data final para filtro de transa��es (usado na consulta SELECT).

    ERROS POSS�VEIS:
        - Caso @Id_Transacao n�o seja fornecido para opera��es de atualiza��o ou exclus�o, um erro ser� gerado.
        - Caso o @Id_Cliente n�o exista na tabela Clientes, um erro ser� gerado nas opera��es de 'INSERT', 'UPDATE' ou 'DELETE'.
        - Se ocorrer uma falha durante a execu��o de qualquer opera��o, a transa��o ser� revertida (rollback).

    EXEMPLOS DE USO:
        -- Inserir uma nova transa��o:
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'INSERT', @Id_Cliente = 1, 
                                      @Numero_Cartao = '1111111111111111', 
                                      @Valor_Transacao = 100.00, 
                                      @Data_Inicial = '2024-12-31', 
                                      @Descricao = 'Compra no supermercado';

        -- Alterar uma transa��o existente:
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'UPDATE', @Id_Transacao = 1, 
                                      @Valor_Transacao = 170.85, 
                                      @Descricao = 'Compra no ARAGUAIA shopping';

        -- Excluir uma transa��o (soft delete):
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'DELETE', @Id_Transacao = 1;

        -- Consultar transa��es:
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Id_Cliente = 1';
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Numero_Cartao ='1111111111111111'
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Data_Inicial = '2024-01-01', @Data_Final = '2024-12-31';
		EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Valor_Transacao = 100.00;
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Descricao = '%SUPERMERCADO%';
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Numero_Cartao ='2222222222222222', @Descricao = '%SUPERMERCADO%';
        EXEC sp_CartaoTransacoes_CRUD @Acao = 'SELECT', @Numero_Cartao ='1111111111111111', @Data_Inicial = '2024-01-01', @Data_Final = '2025-12-31';
 ====================================================================================================================================================== */
CREATE PROCEDURE sp_CartaoTransacoes_CRUD
    @Acao VARCHAR(7), -- A��es (INSERT, UPDATE, DELETE, SELECT)
    @Id_Transacao INT = NULL, 
    @Id_Cliente INT = NULL,
    @Numero_Cartao VARCHAR(16) = NULL,
    @Valor_Transacao DECIMAL(12, 2) = NULL,
    @Data_Inicial DATETIME = NULL, -- Par�metro de data inicial
    @Data_Final DATETIME = NULL,   -- Par�metro de data final
    @Descricao VARCHAR(100) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Erro INT = 0;

    BEGIN TRY
        -- Inicia transa��o no banco de dados
        BEGIN TRANSACTION;

        -- Insere nova transa��o de cart�o (A��o: INSERT)
		-- Para esta a��o cpnsiderar o @Data_Inicial como Data_Transa��o
        IF @Acao = 'INSERT' 
        BEGIN
			-- Antes de agir verifica se a transa��o existe
			IF NOT EXISTS (SELECT 1 FROM Clientes WHERE Numero_Cartao = @Numero_Cartao)
            BEGIN
                THROW 50000, 'N�mero de cart�o n�o encontrado.', 1;
            END;

            INSERT INTO Cartao_Transacoes (Id_Cliente, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao)
            VALUES (@Id_Cliente, @Numero_Cartao, @Valor_Transacao, @Data_Inicial, @Descricao);
        END;

        -- Altera transa��o de cart�o (A��o: UPDATE)
		-- Para esta a��o considerar o param�tro @Data_Inicial como Data_Transa��o
        IF @Acao = 'UPDATE' 
        BEGIN
            IF @Id_Transacao IS NULL
            BEGIN
                SET @Erro = 1;
                THROW 50000, 'Id_Transacao inv�lido, campo obrigat�rio para atualiza��o.', 1;
            END;

			-- Antes de agir verifica se a transa��o existe
			IF NOT EXISTS (SELECT 1 FROM Cartao_Transacoes WHERE Id_Transacao = @Id_Transacao) 
            BEGIN
                THROW 50000, 'Transa��o n�o encontrada para atualiza��o.', 1;
            END;

            UPDATE Cartao_Transacoes
            SET
                Numero_Cartao = COALESCE(@Numero_Cartao, Numero_Cartao),
                Valor_Transacao = COALESCE(@Valor_Transacao, Valor_Transacao),
                Data_Transacao = COALESCE(@Data_Inicial, Data_Transacao),
                Descricao = COALESCE(@Descricao, Descricao)
            WHERE Id_Transacao = @Id_Transacao;
        END;

        -- Soft Delete da transa��o de cart�o - Marca a transa��o como inativa 'I' (A��o: DELETE)
        IF @Acao = 'DELETE' 
        BEGIN 
            IF @Id_Transacao IS NULL
            BEGIN
                SET @Erro = 1;
                THROW 50000, 'Id_Transacao inv�lido, campo obrigat�rio para exclus�o.', 1;
            END;

			-- Antes de agir verifica se a transa��o existe
            IF NOT EXISTS (SELECT 1 FROM Cartao_Transacoes WHERE Id_Transacao = @Id_Transacao)
            BEGIN
                THROW 50000, 'Transa��o n�o encontrada para exclus�o.', 1;
            END;

            UPDATE Cartao_Transacoes
            SET Status = 'I'
            WHERE Id_Transacao = @Id_Transacao;
        END;

        -- Consulta transa��es por diversos par�metros (A��o: SELECT)
        IF @Acao = 'SELECT' 
        BEGIN
            SELECT Id_Transacao, Id_Cliente, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status
            FROM Cartao_Transacoes
            WHERE (@Id_Cliente IS NULL OR Id_Cliente = @Id_Cliente)
              AND (@Numero_Cartao IS NULL OR Numero_Cartao = @Numero_Cartao)
              AND (@Valor_Transacao IS NULL OR Valor_Transacao = @Valor_Transacao)
              AND (@Data_Inicial IS NULL AND @Data_Final IS NULL OR Data_Transacao BETWEEN @Data_Inicial AND @Data_Final) -- Per�odo entre duas datas
              AND Status = 'A' -- S� transa��es ativas
        END;

        -- Se tudo correr bem, comita a transa��o no banco de dados
        COMMIT TRANSACTION;
    END TRY

    BEGIN CATCH
        -- Em caso de erro, realiza o rollback da transa��o no banco de dados
        ROLLBACK TRANSACTION;
        SELECT ERROR_MESSAGE() AS ErrorMessage;
    END CATCH
END;
