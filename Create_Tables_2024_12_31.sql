/*========================================================================================================================
    AUTOR: Guilherme Neto
    DATA DE CRIA��O: 31/12/2024
    DESCRI��O: 
    Este script cria duas tabelas, uma chamada 'Clientes' e a outra 'Cartao_Transacoes'e faz o relacionamento necess�rio entre elas. 
	Al�m disso, ele aprimora a performance ao adicionar �ndices nas colunas que devem ser mais usadas nas cl�usulas WHERE, JOIN ou ORDER BY. 
	O comando DELETE pode ocasionar em fragmenta��o nos dados, pois embora exclua registros, 
	o espa�o onde esses dados estavam alocados nas p�ginas n�o � imediatamente otimizado ou reutilizado, levando � fragmenta��o interna. 
	Para solucionar esse problema, eu optei por incluir o campo 'Status' na tabela 'Cartao_Transacoes', possibilitando a implementa��o de 'Soft Delete', 
	que evita essa fragmenta��o e melhora o desempenho geral.
==============================================================================================================================================*/

Use Desenvolvimento;

--	Verifica se a tabela 'Clientes' existe. Se n�o existir, cria a tabela com os campos necess�rios.
--	A tabela 'Clientes' possui o campo 'Id_Cliente' como chave prim�ria e 'Numero_Cartao' como �nico. 

IF OBJECT_ID('Clientes', 'U') IS NULL -- (User Table)
    CREATE TABLE Clientes (
        Id_Cliente INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,   -- Chave prim�ria com incremento autom�tico
        Nome VARCHAR(100) NOT NULL,                   -- Nome do cliente, campo obrigat�rio
        Numero_Cartao VARCHAR(16) NOT NULL,               -- N�mero do cart�o, campo obrigat�rio
        CONSTRAINT UQ_Numero_Cartao UNIQUE (Numero_Cartao)  -- Garante que o n�mero do cart�o seja �nico
    );

-- Verifica se a tabela 'Cartao_Transacoes' existe. Se n�o existir, cria a tabela para armazenar as transa��es.
IF OBJECT_ID('Cartao_Transacoes', 'U') IS NULL
BEGIN
    CREATE TABLE Cartao_Transacoes (
        Id_Transacao BIGINT IDENTITY(1,1) PRIMARY KEY,      -- Chave prim�ria com incremento autom�tico
        Id_Cliente INT NOT NULL,                         -- Chave estrangeira para a tabela 'Clientes'
        Numero_Cartao VARCHAR(16) NOT NULL,              -- Relacionado ao n�mero do cart�o na tabela 'Clientes'
        Valor_Transacao DECIMAL(12, 2) NOT NULL,         -- Valor da transa��o, campo obrigat�rio
        Data_Transacao DATETIME NOT NULL,                -- Data da transa��o, campo obrigat�rio
        Descricao VARCHAR(100) NULL,                     -- Descri��o opcional da transa��o
        Status CHAR(1) DEFAULT 'A' -- A = Ativo, I = Inativo, padr�o 'A' para ativo
    );
END;

GO

-- Cria��o de chave estrangeira entre 'Cartao_Transacoes' e 'Clientes' para o campo 'Id_Cliente'
-- Isso garante que as transa��es s� poder�o ser associadas a um cliente v�lido.
ALTER TABLE Cartao_Transacoes
ADD CONSTRAINT FK_Cartao_Transacoes_Clientes
FOREIGN KEY (Id_Cliente) REFERENCES Clientes(Id_Cliente)

-- Cria��o de chave estrangeira entre 'Cartao_Transacoes' e 'Clientes' para o campo 'Numero_Cartao'
-- Isso garante que as transa��es s� poder�o ser associadas a um n�mero de cart�o v�lido.
ALTER TABLE Cartao_Transacoes
ADD CONSTRAINT FK_Cartao_Transacoes_Numero_Cartao
FOREIGN KEY (Numero_Cartao) REFERENCES Clientes(Numero_Cartao)
GO

-- Cria��o de �ndices para melhorar a performance das consultas na tabela 'Cartao_Transacoes'.
-- O �ndice 'IDX_Id_Cliente' melhora a performance nas consultas filtrando por 'Id_Cliente'.
CREATE INDEX IDX_Id_Cliente ON Cartao_Transacoes(Id_Cliente);

-- O �ndice 'IDX_Numero_Cartao' melhora a performance nas consultas filtrando por 'Numero_Cartao'.
CREATE INDEX IDX_Numero_Cartao ON Cartao_Transacoes(Numero_Cartao);

-- O �ndice 'IDX_Data_Transacao' melhora a performance nas consultas por 'Data_Transacao'.
CREATE INDEX IDX_Data_Transacao ON Cartao_Transacoes(Data_Transacao);

--FIM
GO
--=================================================================================================================================
-- Insere dados para testes
INSERT INTO Clientes VALUES ('GUILHERME ALVES NETO', '1111111111111111'); -- Autoincremento, id_cliente = 1
INSERT INTO Clientes VALUES ('GABRIELA RODRIGUES PEREIRA NETO', '2222222222222222'); -- id-cliente = 2

select * from Clientes
select * from Cartao_Transacoes

