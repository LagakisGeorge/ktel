Public Class Issuer
    Public Property vatNumber As String
    Public Property branch As Integer
    Public Property street As String
    Public Property streetNumber As String
    Public Property postalCode As String
    Public Property city As String
    Public Property country As String
End Class

Public Class Counterpart
    Public Property vatNumber As String
    Public Property branch As Integer
    Public Property name As String
    Public Property street As String
    Public Property streetNumber As String
    Public Property postalCode As String
    Public Property city As String
    Public Property country As String
    Public Property email As String
End Class

Public Class Representative
    Public Property vatNumber As String
    Public Property name As String
End Class

Public Class DeliveryAddress
    Public Property street As String
    Public Property streetNumber As String
    Public Property postalCode As String
    Public Property city As String
    Public Property country As String
End Class

Public Class Budget
    Public Property type As Integer
    Public Property identifier As String
End Class

Public Class PublishDetails
    Public Property contractingAuthorityID As String
    Public Property budget As Budget
    Public Property contractIdentifier As String
End Class

Public Class InvoiceHeader
    Public Property series As String
    Public Property aa As Integer
    Public Property issueDate As String
    Public Property dispatchDate As String
    Public Property invoiceCode As String
    Public Property invoiceType As String
    Public Property invoiceTypeUbl As String
    Public Property currency As String
    Public Property selfPricing As Boolean
    Public Property movePurpose As Integer
    Public Property fuelInvoice As Boolean
End Class

Public Class TaxInfo
    Public Property taxCategory As Integer
    Public Property taxCategoryUbl As String
    Public Property underlyingValue As Integer
    Public Property taxPercent As Integer
End Class

Public Class InvoiceDetail
    Public Property lineNumber As Integer
    Public Property recType As Integer
    Public Property quantity As Integer
    Public Property entityName As String
    Public Property invoiceDetailType As Integer
    Public Property netValue As Integer
    Public Property totalValue As Integer
    Public Property vatCategory As Integer
    Public Property vatCategoryUbl As String
    Public Property vatExemption As Integer
    Public Property vatExemptionUbl As String
    Public Property vatAmount As Integer
    Public Property vatPercent As Integer
    Public Property measurementUnit As Integer
    Public Property measurementUnitUbl As String
    Public Property lineComments As String
    Public Property classificationCategory As String
    Public Property classificationType As String
    Public Property cpvCode As String
    Public Property taxInfo As TaxInfo
End Class

Public Class TaxesTotal
    Public Property taxType As Integer
    Public Property taxCategory As Integer
    Public Property taxCategoryUbl As String
    Public Property underlyingValue As Integer
    Public Property taxAmount As Integer
    Public Property taxPercent As Integer
End Class

Public Class PaymentMethod
    Public Property type As Integer
    Public Property amount As Integer
End Class

Public Class CorrelatedInvoice
    Public Property extSystemId As String
    Public Property mark As Integer
End Class

Public Class InvoiceSummary
    Public Property totalNetValue As Single
    Public Property totalVatAmount As Single
    Public Property totalValue As Single
End Class

Public Class Message
    Public Property type As Integer
    Public Property recipients As String
    Public Property cc As String
    Public Property templateIdentifier As String
End Class

Public Class Invoice
    Public Property issuer As Issuer
    Public Property counterpart As Counterpart
    Public Property representative As Representative
    Public Property deliveryAddress As DeliveryAddress
    Public Property publishType As Integer
    Public Property publishDetails As PublishDetails
    Public Property invoiceHeader As InvoiceHeader
    Public Property invoiceDetails As New List(Of InvoiceDetail)
    Public Property taxesTotals As New List(Of TaxesTotal)
    Public Property paymentMethods As New List(Of PaymentMethod)
    Public Property correlatedInvoices As New List(Of CorrelatedInvoice)
    Public Property invoiceSummary As InvoiceSummary
    Public Property Messages As Message()
End Class


Public Class Source2
    Public Property invoice As Invoice
End Class

Public Class Example
    Public Property externalSystemId As String
    Public Property source As Source2
End Class