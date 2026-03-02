<?php

namespace App\Http\Controllers;

use App\Models\Seksi;
use App\Models\User;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\Hash;

class UserController extends Controller
{
    public function index()
    {
        return view('user.index',[
            'title' => 'User',
            'seksi' => Seksi::all()
        ]);
    }

    public function getDataUser()
    {
        $dt_user = User::with('seksi')->get();
        return datatables()->of($dt_user)
                        ->addColumn('action', function($data){
                            $button = '<a href="javascript:void(0)"  data-id="'.$data->id.'" class="edit_user btn btn-info btn-xs edit-post" data-toggle="modal" data-target="#modal_edit_user" ><i class="fas fa-edit"></i> Edit</a>';
                            $button .= '&nbsp;&nbsp;';   
                            return $button;
                        })
                        ->rawColumns(['action'])        
                        ->addIndexColumn()
                        ->make(true);
    }

    public function addUser()
    {
        $validator = request()->validate([
            'name' => ['required', 'string', 'max:255'],
            'username' => ['required', 'string', 'max:255', 'unique:users'],
            'password' => ['required', 'string', 'confirmed'],
            'seksi_id' => ['required'],
        ],
        [
            'name.required' => 'Nama tidak boleh kosong',
            'name.string' => 'Nama hanya boleh hufuf dan angka',
            'name.max' => 'Nama maksimal 255 karakter',
            'username.required' => 'Username tidak boleh kosong',
            'username.string' => 'Username hanya boleh hufuf dan angka',
            'username.max' => 'Username maksimal 255 karakter',
            'username.unique' => 'Username yang sama sudah terdaftar',
            'password.required' => 'Password tidak boleh kosong',
            'password.string' => 'Password hanya boleh hufuf dan angka',
            'password.confirmed' => 'Password tidak sama',
            'seksi_id.required' => 'Seksi harus diisi',

        ]
    );
            
            
        // if ($validator->fails())
        // {
        //     return response()->json(['errors'=>$validator->errors()->all()]);
        // }

        User::create([
            'name' => request('name'),
            'username' => request('username'),
            'password' => bcrypt(request('password')),
            'seksi_id' => request('seksi_id'),
        ]);

        return response()->json(['success'=>'Data berhasil di input']);

        
    }

    public function gantiPassword()
    {
        return view('user.ganti_password',[
            'title' => 'Ganti Password'
        ]);
    }

    public function editPassword(Request $request)
    {
        if(!(Hash::check($request->old_password,Auth::user()->password))){
            return redirect(route('gantiPassword'))->with('error','Password sekarang tidak cocok');
        }
        $validator = request()->validate([
            'password' => ['required', 'string', 'confirmed'],
            'old_password' => ['required']
            
        ],
        [
            'password.required' => 'Password tidak boleh kososng',
            'password.string' => 'Password hanya boleh hufuf dan angka',
            'password.confirmed' => 'Password tidak sama',
        ]
        );

        User::where('id',Auth::user()->id)->update([
            'password' => bcrypt($request->password)
            ]);

            return redirect(route('gantiPassword'))->with('success','Password berhasil diganti');
    }

    public function getUser($id)
    {
        $data  = User::where('id',$id)->first();
        return response()->json($data);
    }

    public function editUser(Request $request)
    {
        $edit = User::where('id',$request->id)->update([
            'name' => $request->name,
            'seksi_id' => $request->seksi_id,
            ]);
        return response()->json($edit); 
    }

}
