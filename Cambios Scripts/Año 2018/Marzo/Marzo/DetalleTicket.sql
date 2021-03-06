
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DetalleTicket](
	[idDetalle] [int] IDENTITY(1,1) NOT NULL,
	[NumeroFG] [nvarchar](50) NOT NULL,
	[NumeroMesa] [float] NOT NULL,
	[Color] [nvarchar](50) NOT NULL,
	[Talla] [nvarchar](50) NULL,
	[Linea] [int] NULL,
	[Piezas] [int] NULL,
	[Bultos] [int] NULL,
 CONSTRAINT [PK_DetalleTicket] PRIMARY KEY CLUSTERED 
(
	[idDetalle] ASC,
	[NumeroFG] ASC,
	[NumeroMesa] ASC,
	[Color] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
