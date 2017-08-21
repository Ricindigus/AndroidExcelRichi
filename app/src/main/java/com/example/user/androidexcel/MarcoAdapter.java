package com.example.user.androidexcel;

import android.support.v7.widget.RecyclerView;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.TextView;

import java.util.ArrayList;

/**
 * Created by user on 9/08/2017.
 */

public class MarcoAdapter extends RecyclerView.Adapter<MarcoAdapter.ViewHolder>{

    ArrayList<ItemMarco> itemMarcos;

    public MarcoAdapter(ArrayList<ItemMarco> itemMarcos) {
        this.itemMarcos = itemMarcos;
    }

    @Override
    public ViewHolder onCreateViewHolder(ViewGroup parent, int viewType) {
        View view = LayoutInflater.from(parent.getContext()).inflate(R.layout.marco_item,parent,false);
        ViewHolder viewHolder = new ViewHolder(view);
        return viewHolder;
    }

    @Override
    public void onBindViewHolder(ViewHolder holder, int position) {
        holder.txt1.setText(itemMarcos.get(position).getNumero());
        holder.txt2.setText(itemMarcos.get(position).getRuc());
        holder.txt3.setText(itemMarcos.get(position).getRazonSocial());
    }

    @Override
    public int getItemCount() {
        return itemMarcos.size();
    }

    public static class ViewHolder extends RecyclerView.ViewHolder{
        TextView txt1;
        TextView txt2;
        TextView txt3;

        public ViewHolder(View itemView) {
            super(itemView);
            txt1 = (TextView) itemView.findViewById(R.id.txt1);
            txt2 = (TextView) itemView.findViewById(R.id.txt2);
            txt3 = (TextView) itemView.findViewById(R.id.txt3);
        }
    }
}
