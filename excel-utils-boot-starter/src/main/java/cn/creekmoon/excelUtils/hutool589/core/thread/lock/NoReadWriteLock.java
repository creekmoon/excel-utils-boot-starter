package cn.creekmoon.excelUtils.hutool589.core.thread.lock;

import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReadWriteLock;

/**
 * 无锁的读写锁实现
 *
 * @author looly
 * @since 5.8.0
 */
public class NoReadWriteLock implements ReadWriteLock {
	@Override
	public Lock readLock() {
		return cn.creekmoon.excelUtils.hutool589.core.thread.lock.NoLock.INSTANCE;
	}

	@Override
	public Lock writeLock() {
		return NoLock.INSTANCE;
	}
}
