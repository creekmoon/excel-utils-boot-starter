package cn.creekmoon.excelUtils.hutool589.core.convert.impl;

import cn.creekmoon.excelUtils.hutool589.core.convert.AbstractConverter;
import cn.creekmoon.excelUtils.hutool589.core.convert.Convert;

import java.util.concurrent.atomic.AtomicLongArray;

/**
 * {@link AtomicLongArray}转换器
 *
 * @author Looly
 * @since 5.4.5
 */
public class AtomicLongArrayConverter extends AbstractConverter<AtomicLongArray> {
    private static final long serialVersionUID = 1L;

    @Override
    protected AtomicLongArray convertInternal(Object value) {
        return new AtomicLongArray(Convert.convert(long[].class, value));
    }

}
