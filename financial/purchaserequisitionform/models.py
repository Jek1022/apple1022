from __future__ import unicode_literals
from django.core.urlresolvers import reverse
from django.db import models
from django.contrib.auth.models import User
from django.core.validators import MaxValueValidator, MinValueValidator
import datetime


class Prfmain(models.Model):
    PRF_STATUS_CHOICES = (
        ('F', 'For Approval'),
        ('A', 'Approved'),
        ('D', 'Disapproved'),
    )

    STATUS_CHOICES = (
        ('A', 'Active'),
        ('I', 'Inactive'),
        ('C', 'Cancelled'),
        ('O', 'Posted'),
        ('P', 'Printed'),
    )

    RESPONSE_CHOICES = (
        ('A', 'Approved'),
        ('D', 'Disapproved'),
    )

    PRF_TYPE_CHOICES = (
        ('MSRF', 'MSRF'),
        # ('REGULAR', 'REGULAR'), fixed asset items only
        # ('IMPRF', 'IMPORT RF'), # with importing
        # ('REPLENISHMENT', 'REPLENISHMENT'), # manual add of items office supplies
    )

    URGENCY_CHOICES = (
        ('N', 'Normal'),
        ('R', 'Rush'),
    )

    prfnum = models.CharField(max_length=10, unique=True)
    prfdate = models.DateField()
    prftype = models.CharField(max_length=10, choices=PRF_TYPE_CHOICES, default='MSRF')
    inventoryitemtype = models.ForeignKey('inventoryitemtype.Inventoryitemtype', related_name='prfmain_inventoryitemtype_id')
    particulars = models.TextField()
    prfstatus = models.CharField(max_length=1, choices=PRF_STATUS_CHOICES, default='F')
    urgencytype = models.CharField(max_length=1, choices=URGENCY_CHOICES, default='N')
    dateneeded = models.DateField()
    status = models.CharField(max_length=1, choices=STATUS_CHOICES, default='A')
    enterdate = models.DateTimeField(auto_now_add=True)
    modifydate = models.DateTimeField(default=datetime.datetime.now())
    postdate = models.DateTimeField(null=True, blank=True)
    enterby = models.ForeignKey(User, default=1, related_name='prfmain_enter')
    modifyby = models.ForeignKey(User, default=1, related_name='prfmain_modify')
    postby = models.ForeignKey(User, related_name='prfmain_post', null=True, blank=True)
    isdeleted = models.IntegerField(default=0)
    branch = models.ForeignKey('branch.Branch', related_name='prfbranch_id')
    designatedapprover = models.ForeignKey(User, default=2, related_name='prfdesignated_approver')
    actualapprover = models.ForeignKey(User, related_name='prfactual_approver', null=True, blank=True)
    approverresponse = models.CharField(max_length=1, choices=RESPONSE_CHOICES, null=True, blank=True)
    responsedate = models.DateTimeField(null=True, blank=True)
    remarks = models.TextField(max_length=250, null=True, blank=True)
    print_ctr = models.IntegerField(default=0)
    quantity = models.IntegerField()

    # nego vat
    vatable = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    vatexempt = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    vatzerorated = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    grosscost = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    grossamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    vatamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    netamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)

    # uc vat
    uc_vatable = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_vatexempt = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_vatzerorated = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_grosscost = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_grossamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_vatamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_netamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)


    class Meta:
        db_table = 'prfmain'
        ordering = ['-pk']
        permissions = (("view_purchaserequisitionform", "Can view purchaserequisitionform"),
                       ("view_assignprf", "Can view only assigned prf"),
                       ("view_allassignprf", "Can view all prf"),
                       ("can_approveprf", "Can approve prf"),
                       ("can_disapproveprf", "Can disapprove prf"),)

    def get_absolute_url(self):
        return reverse('purchaserequisitionform:detail', kwargs={'pk': self.pk})

    def __str__(self):
        return self.prfnum

    def __unicode__(self):
        return self.prfnum


class Rfprftransaction(models.Model):
    STATUS_CHOICES = (
        ('A', 'Active'),
        ('I', 'Inactive'),
        ('C', 'Cancelled'),
        ('O', 'Posted'),
        ('P', 'Printed'),
    )

    rfmain = models.ForeignKey('requisitionform.Rfmain', related_name='rfmain_rfprftransaction')
    rfdetail = models.ForeignKey('requisitionform.Rfdetail', related_name='rfdetail_rfprftransaction')
    prfmain = models.ForeignKey('purchaserequisitionform.Prfmain', related_name='prfmain_rfprftransaction')
    prfdetail = models.ForeignKey('purchaserequisitionform.Prfdetail', related_name='prfdetail_rfprftransaction')
    prfquantity = models.IntegerField()
    status = models.CharField(max_length=1, choices=STATUS_CHOICES, default='A')

    class Meta:
        db_table = 'rfprftransaction'
        ordering = ['-pk']

    def get_absolute_url(self):
        return reverse('purchaserequisitionform:detail', kwargs={'pk': self.pk})

    def __str__(self):
        return self.prfnum

    def __unicode__(self):
        return self.prfnum


class Prfdetail(models.Model):
    STATUS_CHOICES = (
        ('A', 'Active'),
        ('I', 'Inactive'),
        ('C', 'Cancelled'),
        ('O', 'Posted'),
        ('P', 'Printed'),
    )

    prfmain = models.ForeignKey('purchaserequisitionform.Prfmain', related_name='prfmain_id', null=True, blank=True)
    rfmain = models.ForeignKey('requisitionform.Rfmain', related_name='rfmain_prfdetail', null=True, blank=True)
    rfdetail = models.ForeignKey('requisitionform.Rfdetail', related_name='rfdetail_prfdetail', null=True, blank=True)
    invitem = models.ForeignKey('inventoryitem.Inventoryitem', related_name='prfdetail_invitem_id')
    currency = models.ForeignKey('currency.Currency', related_name='prfdetail_currency')
    amount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    department_code = models.CharField(max_length=25)
    department_name = models.CharField(max_length=250)
    department = models.ForeignKey('department.Department', related_name='prfdetail_department_id')
    invitem_code = models.CharField(max_length=25)
    invitem_name = models.CharField(max_length=250)
    invitem_unitofmeasure = models.ForeignKey('unitofmeasure.Unitofmeasure', related_name='prfdetail_unitofmeasure_id')
    invitem_unitofmeasure_code = models.CharField(max_length=50)
    item_counter = models.IntegerField()
    quantity = models.IntegerField()
    remarks = models.TextField(max_length=250, null=True, blank=True)
    status = models.CharField(max_length=1, choices=STATUS_CHOICES, default='A')
    enterby = models.ForeignKey(User, default=1, related_name='prfdetail_enter')
    enterdate = models.DateTimeField(auto_now_add=True)
    modifyby = models.ForeignKey(User, default=1, related_name='prfdetail_modify')
    modifydate = models.DateTimeField(default=datetime.datetime.now())
    postby = models.ForeignKey(User, default=1, related_name='prfdetail_post', null=True, blank=True)
    postdate = models.DateTimeField(default=datetime.datetime.now(), null=True, blank=True)
    isdeleted = models.IntegerField(default=0)

    fxrate = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=5, max_digits=18)

    # nego vat
    vatable = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    vatexempt = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    vatzerorated = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    grosscost = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    grossamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    vatamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    netamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)

    # uc vat
    uc_vatable = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_vatexempt = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_vatzerorated = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_grosscost = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_grossamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_vatamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_netamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)

    # cs detail
    csmain = models.ForeignKey('canvasssheet.Csmain', related_name='prfdetail_csmain_id', null=True, blank=True)
    csdetail = models.ForeignKey('canvasssheet.Csdetail', related_name='prfdetail_csdetail_id', null=True, blank=True)
    csnum = models.CharField(max_length=10, null=True, blank=True)
    csdate = models.DateField(null=True, blank=True)

    # supplier
    supplier = models.ForeignKey('supplier.Supplier', related_name='prfdetail_supplier_id', null=True, blank=True)
    suppliercode = models.CharField(max_length=10, null=True, blank=True)
    suppliername = models.CharField(max_length=250, null=True, blank=True)


    class Meta:
        db_table = 'prfdetail'
        ordering = ['-pk']

    def get_absolute_url(self):
        return reverse('purchaserequisitionform:detail', kwargs={'pk': self.pk})

    def __str__(self):
        return str(self.pk) + ' ' + str(self.item_counter) + ' ' + self.invitem_name

    def __unicode__(self):
        return str(self.pk) + ' ' + str(self.item_counter) + ' ' + self.invitem_name


class Prfdetailtemp(models.Model):
    STATUS_CHOICES = (
        ('A', 'Active'),
        ('I', 'Inactive'),
        ('C', 'Cancelled'),
        ('O', 'Posted'),
        ('P', 'Printed'),
    )

    prfmain = models.ForeignKey('purchaserequisitionform.Prfmain', related_name='temp_prfmain_id', null=True, blank=True)
    rfmain = models.ForeignKey('requisitionform.Rfmain', related_name='rfmain_prfdetailtemp', null=True, blank=True)
    rfdetail = models.ForeignKey('requisitionform.Rfdetail', related_name='temp_rfdetail_id', null=True, blank=True)
    prfdetail = models.ForeignKey('purchaserequisitionform.Prfdetail', related_name='temp_prfdetail_id', null=True, blank=True)
    invitem = models.ForeignKey('inventoryitem.Inventoryitem', related_name='temp_prfdetail_invitem_id')
    currency = models.ForeignKey('currency.Currency', related_name='prfdetailtemp_currency')
    amount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    department_code = models.CharField(max_length=25)
    department_name = models.CharField(max_length=250)
    department = models.ForeignKey('department.Department', related_name='prfdetailtemp_department_id')
    invitem_code = models.CharField(max_length=25)
    invitem_name = models.CharField(max_length=250)
    invitem_unitofmeasure = models.ForeignKey('unitofmeasure.Unitofmeasure', related_name='prfdetailtemp_unitofmeasure_id')
    invitem_unitofmeasure_code = models.CharField(max_length=50)
    item_counter = models.IntegerField()
    quantity = models.IntegerField()
    remarks = models.TextField(max_length=250, null=True, blank=True)
    status = models.CharField(max_length=1, choices=STATUS_CHOICES, default='A')
    enterby = models.ForeignKey(User, default=1, related_name='prfdetailtemp_enter')
    enterdate = models.DateTimeField(auto_now_add=True)
    modifyby = models.ForeignKey(User, default=1, related_name='prfdetailtemp_modify')
    modifydate = models.DateTimeField(default=datetime.datetime.now())
    postby = models.ForeignKey(User, related_name='prfdetailtemp_post', null=True, blank=True)
    postdate = models.DateTimeField(null=True, blank=True)
    isdeleted = models.IntegerField(default=0)
    secretkey = models.CharField(max_length=255)

    fxrate = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=5, max_digits=18)

    # nego vat
    vatable = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    vatexempt = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    vatzerorated = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    grosscost = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    grossamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    vatamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    netamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)

    # uc vat
    uc_vatable = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_vatexempt = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_vatzerorated = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_grosscost = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_grossamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_vatamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)
    uc_netamount = models.DecimalField(default=0.00, null=True, blank=True, decimal_places=2, max_digits=18)

    # cs detail
    csmain = models.ForeignKey('canvasssheet.Csmain', related_name='prfdetailtemp_csmain_id', null=True, blank=True)
    csdetail = models.ForeignKey('canvasssheet.Csdetail', related_name='prfdetailtemp_csdetail_id', null=True, blank=True)
    csnum = models.CharField(max_length=10, null=True, blank=True)
    csdate = models.DateField(null=True, blank=True)

    # supplier
    supplier = models.ForeignKey('supplier.Supplier', related_name='prfdetailtemp_supplier_id', null=True, blank=True)
    suppliercode = models.CharField(max_length=10, null=True, blank=True)
    suppliername = models.CharField(max_length=250, null=True, blank=True)

    class Meta:
        db_table = 'prfdetailtemp'
        ordering = ['-pk']

    def get_absolute_url(self):
        return reverse('purchaserequisitionform:detail', kwargs={'pk': self.pk})

    def __str__(self):
        return str(self.pk) + ' ' + str(self.item_counter) + ' ' + self.invitem_name

    def __unicode__(self):
        return str(self.pk) + ' ' + str(self.item_counter) + ' ' + self.invitem_name
