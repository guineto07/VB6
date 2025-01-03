
-- Criação da tabela Clientes
CREATE TABLE Clientes (
    Id_Cliente INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,  -- Chave primária com identidade e autoincremento 
    Nome NVARCHAR(100) NOT NULL,                            
    Numero_Cartao CHAR(16) NOT NULL,                             
    CONSTRAINT UQ_Numero_Cartao UNIQUE (Numero_Cartao)      -- Garantir unicidade do número do cartão
);

-- Índice não-clustered para acelerar buscas por Numero_Cartao
CREATE NONCLUSTERED INDEX IX_Numero_Cartao ON Clientes (Numero_Cartao);

GO	  
	INSERT INTO Clientes(Nome, Numero_Cartao ) VALUES ('GUILHERME ALVES NETO','1111222233334444');    
	INSERT INTO Clientes(Nome, Numero_Cartao ) VALUES ('GABRIELA RODRIGUES PEREIRA NETO','9999888877776666');
-- Criação da tabela Cartao_Transacoes
CREATE TABLE Cartao_Transacoes(
    Id_Transacao INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,  -- Chave primária com autoincremento
    Id_Cliente INT NOT NULL,                                -- Relacionamento com tabela Cliente
    Numero_Cartao CHAR(16) NOT NULL,                                
    Valor_Transacao DECIMAL(12,2) NOT NULL,                    
    Data_Transacao DATETIME NOT NULL,                          
    Descricao NVARCHAR(55) NOT NULL,                            
    -- Chave estrangeira para garantir integridade referencial
    CONSTRAINT FK_Id_Cliente FOREIGN KEY (Id_Cliente) REFERENCES Clientes(Id_Cliente)
);

-- Índices não-clustered para otimizar buscas em Cartao_Transacoes
CREATE NONCLUSTERED INDEX idx_Cartao_Transacoes_Id_Cliente ON Cartao_Transacoes(Id_Cliente);
CREATE NONCLUSTERED INDEX idx_Cartao_Transacoes_Numero_Cartao ON Cartao_Transacoes(Numero_Cartao);

GO

-- Procedure para transações em um período específico
CREATE PROCEDURE sp_TransacaoPorPeriodo
    @Data_Inicial DATE,
    @Data_Final DATE
AS
BEGIN
    SET NOCOUNT ON;  -- Reduz sobrecarga de rede, desativando o retorno de contagem de linhas

    SELECT 
        Numero_Cartao,
        SUM(Valor_Transacao) AS Valor_Total,
        COUNT(*) AS Qtde_Transacoes
    FROM 
        Cartao_Transacoes
    WHERE 
        Data_Transacao BETWEEN @Data_Inicial AND @Data_Final
    GROUP BY 
        Numero_Cartao;
END;

GO

-- Função para determinar a categoria da transação com base no valor
CREATE FUNCTION dbo.fn_CategoriaTransacao (@Valor_Transacao DECIMAL(12,2))
RETURNS VARCHAR(50)
AS
BEGIN
    RETURN CASE 
            WHEN @Valor_Transacao > 1000 THEN 'Alta'
            WHEN @Valor_Transacao BETWEEN 500 AND 1000 THEN 'Média'
            ELSE 'Baixa'
           END;
END;

GO

-- View que combina transações e categorias
CREATE VIEW dbo.vw_TransacoesComCategoria
AS
SELECT 
    Clientes.Id_Cliente,                    
    Clientes.Numero_Cartao,                   
    Cartao_Transacoes.Valor_Transacao,                 
    Cartao_Transacoes.Data_Transacao,               
    dbo.fn_CategoriaTransacao(Cartao_Transacoes.Valor_Transacao) AS Categoria  
FROM 
    Cartao_Transacoes
INNER JOIN 
    Clientes
    ON Cartao_Transacoes.Id_Cliente = Clientes.Id_Cliente; 

