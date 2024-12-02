<?php

namespace mradang\LaravelExcel\Test;

use Illuminate\Database\Eloquent\Model;

class User extends Model
{
    protected $fillable = ['name', 'age', 'date1', 'date2'];
}
