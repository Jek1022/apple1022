# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-11-08 09:47
from __future__ import unicode_literals

import datetime
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('chartofaccount', '0026_auto_20171108_1747'),
        ('inventoryitemclass', '0004_auto_20170322_1356'),
    ]

    operations = [
        migrations.AddField(
            model_name='inventoryitemclass',
            name='depreciationchartofaccount',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='invclass_depreciationchartofaccount_id', to='chartofaccount.Chartofaccount'),
        ),
        migrations.AlterField(
            model_name='inventoryitemclass',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 11, 8, 17, 47, 32, 287000)),
        ),
    ]
