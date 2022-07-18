<?php

use PhpOffice\PhpSpreadsheet\Calculation\Information\ExcelError;

return [
    'Case-sensitive Offset 1' => [
        "'s red hood",
        [
            "Red riding hood's red hood",
            'hood',
        ],
    ],
    'Case-sensitive Offset 2' => [
        '',
        [
            "Red riding hood's red hood",
            'hood',
            2,
        ],
    ],
    'Case-sensitive Offset -1' => [
        '',
        [
            "Red riding hood's red hood",
            'hood',
            -1,
        ],
    ],
    'Case-sensitive Offset -2' => [
        "'s red hood",
        [
            "Red riding hood's red hood",
            'hood',
            -2,
        ],
    ],
    'Case-sensitive Offset 3' => [
        ExcelError::NA(),
        [
            "Red riding hood's red hood",
            'hood',
            3,
        ],
    ],
    'Case-sensitive Offset -3' => [
        ExcelError::NA(),
        [
            "Red riding hood's red hood",
            'hood',
            -3,
        ],
    ],
    'Case-sensitive - No Match' => [
        ExcelError::NA(),
        [
            "Red riding hood's red hood",
            'HOOD',
        ],
    ],
    'Case-insensitive Offset 1' => [
        "'s red hood",
        [
            "Red riding hood's red hood",
            'HOOD',
            1,
            1,
        ],
    ],
    'Case-insensitive Offset 2' => [
        '',
        [
            "Red riding hood's red hood",
            'HOOD',
            2,
            1,
        ],
    ],
    'Offset 0' => [
        ExcelError::VALUE(),
        [
            "Red riding hood's red hood",
            'hood',
            0,
        ],
    ],
    'Empty match positive' => [
        "Red riding hood's red hood",
        [
            "Red riding hood's red hood",
            '',
        ],
    ],
    'Empty match negative' => [
        '',
        [
            "Red riding hood's red hood",
            '',
            -1,
        ],
    ],
    [
        ExcelError::NA(),
        [
            'Socrates',
            ' ',
            1,
            0,
            0,
        ],
    ],
    [
        '',
        [
            'Socrates',
            ' ',
            1,
            0,
            1,
        ],
    ],
];
