package com.example.user.androidexcel;

/**
 * Created by user on 9/08/2017.
 */

public class ItemMarco {
    String numero;
    String ruc;
    String razonSocial;

    public ItemMarco() {
    }

    public ItemMarco(String numero, String ruc, String razonSocial) {
        this.numero = numero;
        this.ruc = ruc;
        this.razonSocial = razonSocial;
    }

    public String getNumero() {
        return numero;
    }

    public void setNumero(String numero) {
        this.numero = numero;
    }

    public String getRuc() {
        return ruc;
    }

    public void setRuc(String ruc) {
        this.ruc = ruc;
    }

    public String getRazonSocial() {
        return razonSocial;
    }

    public void setRazonSocial(String razonSocial) {
        this.razonSocial = razonSocial;
    }
}
