
'������� ��� ��� 
'    [ENTITYUID] [varchar](40) NULL,
'	[ENTITYMARK] [varchar](43) NULL,
'	[ENTITY] [int] NULL,
'	[AADEKAU] [float] NULL,
'	[AADEFPA] [float] NULL,
'	[ENTLINEN] [int] NULL,
'	[INCMARK] [nvarchar](43) NULL,


��� ��� => INVyyyyddmmhhmm.xml
�������� ����   INVyyyyddmmhhmm.xml => apantSendInv.xml


 apantSendInv.xml ��������� ��� (ENTITY , ENTITYMARK )
                  ���������� INC.XML (��������� ��� (EntLineN) )

�������� ����   INc.xml => apantIncome.xml


apantIncome.xml  ��������� ��� ( incMARK ) ������������ ( (EntLineN) <=> apantIncome.xml  )
