# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2018-04-24 10:04
from __future__ import unicode_literals

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('subledger', '0008_auto_20180424_1802'),
    ]

    operations = [
        migrations.AlterField(
            model_name='logs_subledger',
            name='document_fxamount',
            field=models.DecimalField(blank=True, decimal_places=2, default=0.0, max_digits=18, null=True),
        ),
        migrations.AlterField(
            model_name='logs_subledger',
            name='document_fxrate',
            field=models.DecimalField(blank=True, decimal_places=2, default=0.0, max_digits=18, null=True),
        ),
        migrations.AlterField(
            model_name='logs_subledger',
            name='fxamount',
            field=models.DecimalField(blank=True, decimal_places=2, default=0.0, max_digits=18, null=True),
        ),
        migrations.AlterField(
            model_name='logs_subledger',
            name='fxrate',
            field=models.DecimalField(blank=True, decimal_places=2, default=0.0, max_digits=18, null=True),
        ),
        migrations.AlterField(
            model_name='subledger',
            name='document_fxamount',
            field=models.DecimalField(blank=True, decimal_places=2, default=0.0, max_digits=18, null=True),
        ),
        migrations.AlterField(
            model_name='subledger',
            name='document_fxrate',
            field=models.DecimalField(blank=True, decimal_places=2, default=0.0, max_digits=18, null=True),
        ),
        migrations.AlterField(
            model_name='subledger',
            name='fxamount',
            field=models.DecimalField(blank=True, decimal_places=2, default=0.0, max_digits=18, null=True),
        ),
        migrations.AlterField(
            model_name='subledger',
            name='fxrate',
            field=models.DecimalField(blank=True, decimal_places=2, default=0.0, max_digits=18, null=True),
        ),
    ]
