package org.unisens.unisens2excel;

import java.util.Comparator;

import org.unisens.Entry;

public class UnisensEntryComparer implements Comparator<org.unisens.Entry>
{

    @Override
    public int compare(Entry entry0, Entry entry1)
    {

        return (entry0.getId().compareTo(entry1.getId()));

    }

}
