# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-11-22 06:28
from __future__ import unicode_literals

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('product', '0013_auto_20171107_2251'),
    ]

    operations = [
        migrations.AlterField(
            model_name='product',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 11, 22, 14, 28, 46, 129000)),
        ),
    ]
