{
  'targets': [
    {
      'target_name': 'node_ole',
      'sources': [
        'cpp/main.cpp',
        'cpp/handler.cpp',
        'cpp/environment.cpp',
		'cpp/comutil.cpp'
      ],
      'dependencies': [
      ],
      "include_dirs" : [
        "<!(node -e \"require('nan')\")"
      ]
    }
  ]
}